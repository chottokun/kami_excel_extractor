"""
Kami Excel Extractor のコアオーケストレーションエンジン。
マルチモーダルLLM、視覚解析、およびロジック解析を統合し、構造化データを生成する。
"""

import asyncio
import base64
import html
import json
import logging
import os
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import aiofiles
import litellm
import yaml

from .converter import ExcelConverter
from .document_generator import DocumentGenerator
from .extractor import MetadataExtractor
from .jsonl_exporter import JsonlExporter
from .rag_converter import ContextualChunkGenerator, JsonToMarkdownConverter
from .schema import ExtractionOptions, ExtractionResult, RagOptions
from .utils import CacheManager

logger = logging.getLogger(__name__)


class KamiExcelExtractor:
    """
    Excelの構造化抽出を統合管理するメインクラス。

    テキスト抽出、視覚的スタイル、ロジック（計算式）、およびマルチモーダルAIによる
    意味論的解析を組み合わせた高度な抽出パイプラインを提供する。
    """

    _RE_JSON_BLOCK = re.compile(r"```json\s*(.*?)\s*```", re.DOTALL)
    _RE_YAML_BLOCK = re.compile(r"```yaml\s*(.*?)\s*```", re.DOTALL)

    def __init__(
        self,
        output_dir: Union[str, Path] = "output",
        base_url: Optional[str] = None,
        api_key: Optional[str] = None,
        timeout: Optional[int] = None,
        litellm_rpm_limit: int = 0,
    ):
        """
        KamiExcelExtractor を初期化する。
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # 環境変数からのフォールバックを強化 (ハードコード排除)
        self.base_url = base_url or os.getenv("LLM_BASE_URL") or os.getenv("OLLAMA_BASE_URL")
        self.timeout = timeout or float(os.getenv("LLM_TIMEOUT") or 600)
        self.litellm_rpm_limit = litellm_rpm_limit or int(os.getenv("LLM_RPM_LIMIT", 0))
        self.api_key = api_key.strip("'\" ") if api_key else (os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY"))

        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir, max_file_size_mb=ExtractionOptions().max_file_size_mb)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)

        # キャッシュ管理の初期化
        self._db_path = self.output_dir / ".cache.db"
        self.cache = CacheManager(self._db_path)

        self._image_cache = {}
        self._image_locks = {}
        self._visual_summary_cache = {}

        self.default_model = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"

    def _make_json_serializable(self, data: Any) -> Any:
        """オブジェクトをJSONシリアライズ可能な形式に再帰的に変換する。"""
        if isinstance(data, dict):
            return {k: self._make_json_serializable(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self._make_json_serializable(i) for i in data]
        elif isinstance(data, (date, datetime)):
            return data.isoformat()
        return data

    async def _encode_image_to_base64_url(self, image_path: Path) -> str:
        """画像をBase64エンコードし、データURL形式で返す。永続キャッシュ対応。"""
        # ハッシュを計算して内容ベースで管理
        img_hash = await self.cache.aget_file_hash(image_path)

        use_cache = getattr(self, "opts", None) is None or self.opts.use_cache

        # 1. メモリキャッシュ
        if use_cache and img_hash in self._image_cache:
            return self._image_cache[img_hash]

        # 2. 永続キャッシュ(SQLite)
        if use_cache:
            cached_url = self.cache.get_image_data_url(img_hash)
            if cached_url:
                self._image_cache[img_hash] = cached_url
                return cached_url

        # 同一ハッシュへの同時エンコードを防ぐためのロック取得
        lock = self._image_locks.setdefault(img_hash, asyncio.Lock())
        async with lock:
            if use_cache and img_hash in self._image_cache:
                return self._image_cache[img_hash]

            async with aiofiles.open(image_path, "rb") as f:
                content = await f.read()

            def _encode(data: bytes) -> str:
                encoded = base64.b64encode(data).decode("utf-8")
                return f"data:image/png;base64,{encoded}"

            result = await asyncio.to_thread(_encode, content)

            # キャッシュ保存
            if use_cache:
                self.cache.set_image_data_url(img_hash, result)
                self._image_cache[img_hash] = result
            return result

    def _resolve_model(self, model: Optional[str] = None) -> str:
        """使用するモデル名を解決する。"""
        return model or self.default_model

    async def _awith_retry(self, func, *args, max_retries: int = 3, initial_delay: float = 2.0, **kwargs):
        """
        指数関数的バックオフを用いて非同期関数をリトライ実行する。
        主にAPIのレートリミット(429)対策。
        """
        last_exception = None
        for attempt in range(max_retries + 1):
            try:
                return await func(*args, **kwargs)
            except Exception as e:
                last_exception = e
                # レートリミット(429)またはサーバーエラー(5xx)の場合にリトライ
                status_code = getattr(e, "status_code", None)
                if attempt < max_retries and (status_code == 429 or status_code is None or 500 <= status_code < 600):
                    delay = initial_delay * (2**attempt)
                    logger.warning(
                        f"API error ({status_code}). Retrying in {delay:.1f}s (Attempt {attempt + 1}/{max_retries})..."
                    )
                    await asyncio.sleep(delay)
                else:
                    break
        raise last_exception

    def _get_semaphore(self) -> asyncio.Semaphore:
        """RPM制限用のセマフォを取得する。制限なし(0)の場合は十分に大きな値を設定。"""
        limit = self.litellm_rpm_limit if self.litellm_rpm_limit > 0 else 1000
        return asyncio.Semaphore(limit)

    def _build_sheet_messages(
        self,
        system_prompt: str,
        sheet_name: str,
        html_content: str,
        image_urls: Optional[List[str]] = None,
        include_logic: bool = False,
    ) -> List[Dict]:
        """LLMへの入力メッセージ（プロンプト）を構築する。"""
        context_instruction = (
            "提供されたHTMLテーブルには、CSSスタイル属性(style)が含まれています。\n"
            "- border属性はデータの区切りや表の構造を強く示唆しています。\n"
            "- background-colorはヘッダーや特定のデータグループを示していることが多いです。\n"
            "- data-coord属性はExcel上の絶対座標を示しており、離れたデータ間の関係推論に役立ちます。\n"
        )

        if include_logic:
            context_instruction += (
                "- data-formula属性はセルの計算式を示しており、=SUM(...) 等は合計値を意味します。\n"
                "- data-unit属性はセルの単位（JPY, PERCENT等）です。\n"
            )

        context_instruction += "これらの情報を活用し、人間が目で見た時と同じ論理構造を復元してください。"

        text_payload = f"対象シート: {sheet_name}\n\n{context_instruction}\n\nデータソース (HTML):\n{html_content}"
        content = [{"type": "text", "text": text_payload}]

        if image_urls:
            for url in image_urls:
                content.append({"type": "image_url", "image_url": {"url": url}})

        content.append(
            {
                "type": "text",
                "text": "解析結果を構造化されたJSONオブジェクトとして出力してください。必ず ```json ブロックを含めてください。",
            }
        )

        return [{"role": "system", "content": system_prompt}, {"role": "user", "content": content}]

    def _parse_llm_response(self, content: str, sheet_name: str) -> Dict:
        """LLMの応答からJSON/YAMLを抽出し、検証して辞書として返す。"""
        json_match = self._RE_JSON_BLOCK.search(content)
        json_str = json_match.group(1) if json_match else ""

        data = None
        raw_str = ""

        if json_str:
            try:
                data = json.loads(json_str)
                raw_str = json_str
            except json.JSONDecodeError as e:
                logger.debug(f"JSON decoding failed for {sheet_name}, falling back to YAML: {e}")

        if data is None:
            yaml_match = self._RE_YAML_BLOCK.search(content)
            yaml_str = yaml_match.group(1) if yaml_match else content
            try:
                data = yaml.safe_load(yaml_str)
                raw_str = yaml_str
            except Exception as e:
                logger.error(f"Failed to parse response for {sheet_name}: {e}")
                return {"error": str(e), "_raw_data": content}

        try:
            # 1. 辞書形式であることを保証
            if not isinstance(data, dict):
                data = {"data": data}

            # 2. 'sheets' キーによるネスティングの解消
            if "sheets" in data and isinstance(data["sheets"], dict) and sheet_name in data["sheets"]:
                data = data["sheets"][sheet_name]

            # 3. 最終的な正規化: Pydanticの extra='ignore' による消失を防ぐため 'data' キーを保証
            if not isinstance(data, dict):
                data = {"data": data}
            elif "data" not in data and "error" not in data:
                data = {"data": data}

            # Pydanticによるスキーマ検証とクレンジング
            # model_validate -> model_dump() により、extra='ignore' 設定に基づき
            # 未知のフィールドが自動的に削除された辞書が得られる。
            validated_obj = ExtractionResult(**data)
            cleaned_data = validated_obj.model_dump(exclude_none=True)

            # 生の応答テキストを保持 (テストおよびデバッグ用)
            cleaned_data["_raw_data"] = raw_str
            return cleaned_data
        except Exception as e:
            logger.error(f"Validation failed for {sheet_name}: {e}")
            return {"error": f"Validation failed: {str(e)}", "_raw_data": raw_str}

    async def _aprocess_chart_data(self, media_item: Dict, model: str, semaphore: Optional[asyncio.Semaphore]) -> Dict:
        """画像がグラフや図表の場合、VLMを用いてその内容を構造化データとして抽出する。"""
        if not media_item.get("filename"):
            return media_item
        image_path = self.output_dir / "media" / media_item["filename"]
        if not image_path.exists():
            return media_item

        async with semaphore if semaphore else asyncio.Lock():
            try:
                image_url = await self._encode_image_to_base64_url(image_path)
                prompt = (
                    "この画像がグラフや図表の場合、その軸ラベル、凡例、およびデータ値を抽出し、"
                    "Markdownのテーブル形式で整理してください。回答の冒頭に [図表データ] と付けて出力してください。"
                )
                messages = [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": image_url}},
                        ],
                    }
                ]

                response = await self._awith_retry(
                    litellm.acompletion,
                    model=model,
                    messages=messages,
                    api_key=self.api_key,
                    base_url=self.base_url,
                    timeout=self.timeout,
                )
                media_item["visual_data"] = response.choices[0].message.content or ""
                return media_item
            except Exception as e:
                logger.warning(f"Failed to extract chart data for {image_path.name}: {e}")
                return media_item

    def _inject_visual_data_to_html(self, html_content: str, media_map: Dict) -> str:
        """抽出された図表データをHTMLテーブル内の該当セルに動的に注入する。"""
        for coord, items in media_map.items():
            insight_html = self._format_visual_insights(coord, items)
            if insight_html:
                html_content = self._inject_insight_to_html(html_content, coord, insight_html)
        return html_content

    def _format_visual_insights(self, coord: str, items: List[Dict]) -> str:
        """指定された座標の図表データをHTML形式にフォーマットする。"""
        insights = [
            f"<div class='visual-insight'>[図表データ({coord})]: {html.escape(i['visual_data'])}</div>"
            for i in items
            if "visual_data" in i
        ]
        return "\n".join(insights)

    def _inject_insight_to_html(self, html_content: str, coord: str, insight_html: str) -> str:
        """HTML内の特定の座標属性を持つ箇所にインサイトを注入する。"""
        for attr_pattern in [f'data-coord="{coord}"', f"data-coord='{coord}'"]:
            if attr_pattern in html_content:
                # セル属性の直後にデータを挿入
                return html_content.replace(attr_pattern, f"{attr_pattern} {insight_html}")
        return html_content

    async def _aprocess_media_summary(
        self, media_item: Dict, model: str, semaphore: Optional[asyncio.Semaphore], use_cache: bool = True
    ) -> Optional[Dict]:
        """メディアアイテムの要約文と図表データの抽出を並列実行する。"""
        if not media_item.get("filename"):
            return media_item
        media_path = self.output_dir / "media" / media_item["filename"]
        if not media_path.exists():
            return None

        # aget_visual_summary と _aprocess_chart_data 内部でそれぞれセマフォを制御
        summary = await self.aget_visual_summary(media_path, model=model, semaphore=semaphore, use_cache=use_cache)
        media_item["visual_summary"] = summary
        await self._aprocess_chart_data(media_item, model, semaphore)
        return media_item

    async def _aextract_single_sheet(
        self,
        sheet_name: str,
        sheet_content: Dict,
        model: str,
        system_prompt: str,
        image_urls: Optional[List[str]],
        semaphore: Optional[asyncio.Semaphore],
        include_logic: bool = False,
        use_cache: bool = True,
    ) -> Tuple[str, Dict]:
        """単一シートの解析。シンプルテーブルの場合はLLMをバイパスする。"""
        if sheet_content.get("is_simple"):
            logger.info(f"Using simple table extraction for: {sheet_name}")
            result = sheet_content.get("structured_data", [])
            return sheet_name, {"data": result if isinstance(result, dict) else {"data": result}, "_raw_data": ""}

        async with semaphore if semaphore else asyncio.Lock():
            logger.info(f"Processing via LLM: {sheet_name} (Images: {len(image_urls) if image_urls else 0})")
            html_content = sheet_content.get("html", "")
            messages = self._build_sheet_messages(
                system_prompt, sheet_name, html_content, image_urls, include_logic=include_logic
            )

            # 1. 永続キャッシュのチェック
            if use_cache:
                cached_result_str = await self.cache.aget_llm_result(model, system_prompt, html_content)
                if cached_result_str:
                    return sheet_name, self._parse_llm_response(cached_result_str, sheet_name)

            max_retries = 1
            last_error = None
            for attempt in range(max_retries + 1):
                if attempt > 0:
                    messages.append(
                        {
                            "role": "assistant",
                            "content": f"エラー内容: {last_error}\n修正して正しいJSONで再出力してください。",
                        }
                    )
                try:
                    fmt = {"type": "json_object"} if any(m in model for m in ["ollama", "gpt", "gemini"]) else None
                    response = await self._awith_retry(
                        litellm.acompletion,
                        model=model,
                        messages=messages,
                        api_key=self.api_key,
                        base_url=self.base_url,
                        timeout=self.timeout,
                        response_format=fmt,
                    )
                    raw_response = response.choices[0].message.content
                    result = self._parse_llm_response(raw_response, sheet_name)
                    if "error" not in result:
                        # 成功時にキャッシュ保存
                        if use_cache:
                            await self.cache.aset_llm_result(model, system_prompt, html_content, raw_response)
                        return sheet_name, result
                    last_error = result["error"]
                except Exception as e:
                    last_error = str(e)

                await asyncio.sleep(2**attempt)
            return sheet_name, {
                "error": f"Failed after {max_retries} retries. Last error: {last_error}",
                "_raw_data": "",
            }

    async def _validate_file_size(self, excel_path: Path, max_file_size_mb: float) -> None:
        """🔒 Security & Resource Fix: ファイルサイズ制限のチェック"""
        stat_result = await asyncio.to_thread(excel_path.stat)
        file_size_mb = stat_result.st_size / (1024 * 1024)
        if file_size_mb > max_file_size_mb:
            raise ValueError(f"File size ({file_size_mb:.1f}MB) exceeds the limit ({max_file_size_mb:.1f}MB).")

    async def _get_raw_extraction_results(
        self, excel_path: Path, include_logic: bool, use_cache: bool
    ) -> Dict[str, Any]:
        """Extractorによる基本解析 (キャッシュ対応)。"""
        file_hash = await self.cache.aget_file_hash(excel_path)

        if use_cache:
            cached_raw = await self.cache.aget_raw_extraction(file_hash, include_logic)
            if cached_raw:
                try:
                    candidate_data = json.loads(cached_raw)
                    if not self._is_any_media_missing(candidate_data):
                        logger.info("Using cached raw extraction results.")
                        return candidate_data
                    else:
                        logger.warning("Cache hit for raw extraction but media files are missing. Re-extracting.")
                except Exception as e:
                    logger.warning(f"Failed to load cached raw extraction: {e}")

        raw_data = await asyncio.to_thread(self.extractor.extract, excel_path, include_logic=include_logic)
        if use_cache:
            await self.cache.aset_raw_extraction(file_hash, include_logic, json.dumps(raw_data))
        return raw_data

    async def _generate_visual_context(
        self,
        excel_path: Path,
        sheets_data: Dict[str, Any],
        use_visual_context: bool,
        dpi: int,
        max_file_size_mb: float,
        include_visual_summaries: bool,
    ) -> Dict[str, List[str]]:
        """シートごとの画像生成 (ページネーション対応)。"""
        sheet_images = {}
        if use_visual_context:
            for sheet_name in sheets_data.keys():
                try:
                    self.converter.dpi = dpi
                    self.converter.max_file_size_mb = max_file_size_mb
                    # シート個別に画像を生成
                    png_paths = await asyncio.to_thread(self.converter.convert, excel_path, sheet_name=sheet_name)
                    if isinstance(png_paths, Path):
                        png_paths = [png_paths]

                    image_urls = []
                    for p in png_paths:
                        url = await self._encode_image_to_base64_url(p)
                        image_urls.append(url)
                    sheet_images[sheet_name] = image_urls
                except Exception as e:
                    logger.warning(f"Visual context generation failed for sheet '{sheet_name}': {e}")
                    sheet_images[sheet_name] = []

        # 全体概要用の画像 (Summary用)
        if include_visual_summaries:
            try:
                # シート指定なしで全体概要PDFから1枚目を生成
                self.converter.max_file_size_mb = max_file_size_mb
                png_path = await asyncio.to_thread(self.converter.convert, excel_path)
                if isinstance(png_path, list):
                    png_path = png_path[0]
                await self._encode_image_to_base64_url(png_path)
            except Exception as e:
                logger.warning(f"Overall summary image generation failed: {e}")

        return sheet_images

    async def _process_media_summaries(
        self,
        sheets_data: Dict[str, Any],
        model: str,
        semaphore: Optional[asyncio.Semaphore],
        include_visual_summaries: bool,
        use_cache: bool,
    ) -> List[Dict]:
        """メディア（図表・グラフ）の個別解析。"""
        all_media = []
        if include_visual_summaries:
            unique_media = self._get_unique_media(sheets_data)
            media_tasks = [
                self._aprocess_media_summary(m, model, semaphore, use_cache=use_cache) for m in unique_media.values()
            ]
            if media_tasks:
                media_results = await asyncio.gather(*media_tasks)
                all_media = [m for m in media_results if m]
                self._sync_media_results_to_metadata(all_media, sheets_data)
        return all_media

    def _inject_visual_insights(self, sheets_data: Dict[str, Any]) -> None:
        """図表データの注入。"""
        for sheet_name, sheet_info in sheets_data.items():
            if "media_map" in sheet_info:
                sheet_info["html"] = self._inject_visual_data_to_html(sheet_info["html"], sheet_info["media_map"])
                # 🐞 Bug Fix: 小文字のクラス名 "visual-insight" を正しくチェックする
                if "visual-insight" in sheet_info["html"]:
                    logger.warning(f"Injected visual insights for: {sheet_name}")

    async def _perform_llm_sheet_extraction(
        self,
        sheets_data: Dict[str, Any],
        sheet_images: Dict[str, List[str]],
        model: str,
        semaphore: Optional[asyncio.Semaphore],
        system_prompt: Optional[str],
        include_logic: bool,
        use_cache: bool,
    ) -> Dict[str, Dict]:
        """各シートのLLM解析。"""
        sys_prompt = (
            system_prompt or "あなたはExcel構造化の専門家です。HTMLデータを統合し、意味論的なJSONを出力してください。"
        )

        tasks = [
            self._aextract_single_sheet(
                n,
                c,
                model,
                sys_prompt,
                sheet_images.get(n),
                semaphore,
                include_logic=include_logic,
                use_cache=use_cache,
            )
            for n, c in sheets_data.items()
        ]
        results = await asyncio.gather(*tasks)
        return {name: res for name, res in results}

    async def aextract_structured_data(
        self, excel_path: Union[str, Path], options: Optional[ExtractionOptions] = None
    ) -> Dict:
        """Excelを解析して構造化データを取得する (非同期エントリーポイント)。"""
        opts = options or ExtractionOptions()
        model = self._resolve_model(opts.model)
        semaphore = self._get_semaphore()
        excel_path = Path(excel_path)

        # キャッシュ無効化の反映
        if not opts.use_cache:
            logger.info("Cache disabled for this run. Clearing memory cache.")
            self._image_cache = {}
            self._visual_summary_cache = {}

        # 1. バリデーションと基本解析
        await self._validate_file_size(excel_path, opts.max_file_size_mb)
        logger.info(f"Starting extraction for {excel_path.name} (Logic: {opts.include_logic})")
        raw_data = await self._get_raw_extraction_results(
            excel_path, include_logic=opts.include_logic, use_cache=opts.use_cache
        )
        sheets_data = raw_data.get("sheets", {})

        # 2. 視覚的コンテキストの生成
        sheet_images = await self._generate_visual_context(
            excel_path,
            sheets_data,
            use_visual_context=opts.use_visual_context,
            dpi=opts.dpi,
            max_file_size_mb=opts.max_file_size_mb,
            include_visual_summaries=opts.include_visual_summaries,
        )

        # 3. メディア要約の作成とインサイト注入
        all_media = await self._process_media_summaries(
            sheets_data,
            model,
            semaphore,
            include_visual_summaries=opts.include_visual_summaries,
            use_cache=opts.use_cache,
        )
        self._inject_visual_insights(sheets_data)

        # 4. LLMによる構造化抽出の実行
        structured_sheets = await self._perform_llm_sheet_extraction(
            sheets_data,
            sheet_images,
            model,
            semaphore,
            system_prompt=opts.system_prompt,
            include_logic=opts.include_logic,
            use_cache=opts.use_cache,
        )

        # 5. 結果の統合
        final_data = {"sheets": structured_sheets}
        if all_media:
            final_data["media"] = all_media
            self._attach_media_to_sheets(all_media, structured_sheets)

        return self._make_json_serializable(final_data)

    def _get_unique_media(self, sheets_data: Dict[str, Any]) -> Dict[str, Dict]:
        """重複メディア（同一ファイル名）を排除して収集する。"""
        unique_media = {}
        for s in sheets_data.values():
            for m in s.get("media", []):
                if filename := m.get("filename"):
                    if filename not in unique_media:
                        unique_media[filename] = m
        return unique_media

    def _sync_media_results_to_metadata(self, all_media: List[Dict], sheets_data: Dict[str, Any]) -> None:
        """抽出されたメディアの結果をメタデータに同期する。"""
        # 1. ファイル名からターゲットへのマップを作成 (O(1) lookup で高速化)
        filename_to_targets = {}
        for s_info in sheets_data.values():
            for mapped_list in s_info.get("media_map", {}).values():
                for mapped_m in mapped_list:
                    if fname := mapped_m.get("filename"):
                        filename_to_targets.setdefault(fname, []).append(mapped_m)

        # 2. メディアの結果を全ターゲットに反映
        for m in all_media:
            filename = m.get("filename")
            if not filename:
                continue

            visual_data = m.get("visual_data")
            visual_summary = m.get("visual_summary")

            for target in filename_to_targets.get(filename, []):
                if visual_data:
                    target["visual_data"] = visual_data
                if visual_summary:
                    target["visual_summary"] = visual_summary

    def _extract_sheet_name_from_filename(self, filename: str) -> str:
        """メディアファイル名からシート名を推測する (例: Sheet1_img_A1_0.png -> Sheet1)。"""
        if not filename:
            return ""
        return filename.split("_img_")[0] if "_img_" in filename else filename.split("_")[0]

    def _attach_media_to_sheets(self, all_media: list, structured_sheets: dict) -> None:
        """メディア情報をシートデータに紐付ける。"""
        for m in all_media:
            sheet_name = self._extract_sheet_name_from_filename(m.get("filename", ""))
            if sheet_name in structured_sheets:
                structured_sheets[sheet_name].setdefault("media", []).append(m)

    def _is_any_media_missing(self, candidate_data: Dict) -> bool:
        """キャッシュされたデータ内のメディアファイルが実在するか確認する。"""
        for s_data in candidate_data.get("sheets", {}).values():
            for m_item in s_data.get("media", []):
                if fname := m_item.get("filename"):
                    if not (self.output_dir / "media" / fname).exists():
                        return True
        return False

    async def aget_visual_summary(
        self,
        image_path: Path,
        model: Optional[str] = None,
        semaphore: Optional[asyncio.Semaphore] = None,
        use_cache: bool = True,
    ) -> str:
        """画像の視覚的要約を生成する。永続キャッシュ対応。"""
        # 🔒 Security Fix: ファイルサイズ制限のチェック
        stat_result = await asyncio.to_thread(image_path.stat)
        if stat_result.st_size > 20 * 1024 * 1024:  # 20MB limit
            logger.warning(f"Image file too large for visual summary: {image_path.name} ({stat_result.st_size} bytes)")
            return "[画像が大きすぎるため、要約をスキップしました]"

        model = model or self.default_model
        img_hash = await self.cache.aget_file_hash(image_path)
        prompt = "この画像の内容を詳細に説明してください。[画像概要] と付けて出力してください。"

        # 1. メモリキャッシュ
        cache_key = (model, img_hash)
        if use_cache and cache_key in self._visual_summary_cache:
            return self._visual_summary_cache[cache_key]

        # 2. 永続キャッシュ(SQLite)
        if use_cache:
            cached_summary = await self.cache.aget_vlm_result(model, prompt, img_hash)
            if cached_summary:
                self._visual_summary_cache[cache_key] = cached_summary
                return cached_summary

        async with semaphore if semaphore else asyncio.Lock():
            try:
                image_url = await self._encode_image_to_base64_url(image_path)

                messages = [
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": image_url}},
                        ],
                    }
                ]
                response = await self._awith_retry(
                    litellm.acompletion,
                    model=model,
                    messages=messages,
                    api_key=self.api_key,
                    base_url=self.base_url,
                    timeout=self.timeout,
                )
                content = response.choices[0].message.content or ""
                if content:
                    # キャッシュ保存
                    if use_cache:
                        await self.cache.aset_vlm_result(model, prompt, img_hash, content)
                        self._visual_summary_cache[cache_key] = content
                return content
            except Exception as e:
                logger.error(f"Visual summary failed: {e}")
                return "[画像概要] 解析失敗。"

    def extract_structured_data(
        self, excel_path: Union[str, Path], options: Optional[ExtractionOptions] = None
    ) -> Dict:
        return asyncio.run(self.aextract_structured_data(excel_path, options=options))

    def extract_rag_chunks(
        self, excel_path: Union[str, Path], options: Optional[RagOptions] = None
    ) -> Tuple[Dict, Dict]:
        """Excelを解析し、RAG用のMarkdownチャンクと構造化データを同時に生成する。"""
        return asyncio.run(self.aextract_rag_chunks(excel_path, options=options))

    async def aextract_rag_chunks(
        self, excel_path: Union[str, Path], options: Optional[RagOptions] = None
    ) -> Tuple[Dict, Dict]:
        opts = options or RagOptions()
        excel_path = Path(excel_path)
        extract_opts = ExtractionOptions(
            model=opts.model,
            system_prompt=opts.system_prompt,
            include_visual_summaries=True,
            use_visual_context=opts.use_visual_context,
            include_logic=opts.include_logic,
            max_file_size_mb=opts.max_file_size_mb,
            use_cache=opts.use_cache,
        )

        structured_data = await self.aextract_structured_data(excel_path, options=extract_opts)

        # Excel原本データを取得 (キャッシュ経由なので高速)
        raw_data = await self._get_raw_extraction_results(
            excel_path, include_logic=opts.include_logic, use_cache=opts.use_cache
        )

        # RAG出力用ディレクトリの作成
        rag_dir = self.output_dir / f"{excel_path.stem}_rag"
        await asyncio.to_thread(rag_dir.mkdir, parents=True, exist_ok=True)

        sheet_results = {}
        all_jsonl_chunks = []

        for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
            raw_sheet_data = raw_data.get("sheets", {}).get(sheet_name)

            # 新しい ContextualChunkGenerator の使用
            generator = ContextualChunkGenerator(options=opts)
            chunks = generator.generate_chunks(
                sheet_name=sheet_name,
                structured_content={"sheets": {sheet_name: sheet_data}},
                raw_sheet_data=raw_sheet_data,
                source_file=excel_path.name,
            )

            # 後方互換性のためのプレーンなMarkdownテキスト生成
            markdown_text = self.rag_converter.convert({"sheets": {sheet_name: sheet_data}})

            sheet_results[sheet_name] = {"chunks": chunks, "markdown": markdown_text, "structured": sheet_data}

            # ファイルへの書き出し
            if opts.output_format == "yaml_frontmatter":
                for idx, chunk in enumerate(chunks, 1):
                    chunk_file = rag_dir / f"{sheet_name}_chunk_{idx}.md"

                    def _write_chunk(p, c):
                        with open(p, "w", encoding="utf-8") as f:
                            f.write(c)

                    await asyncio.to_thread(_write_chunk, chunk_file, chunk["content"])
            elif opts.output_format == "markdown":
                sheet_file = rag_dir / f"{sheet_name}.md"

                def _write_sheet(p, c):
                    with open(p, "w", encoding="utf-8") as f:
                        f.write(c)

                await asyncio.to_thread(_write_sheet, sheet_file, markdown_text)
            elif opts.output_format == "jsonl":
                all_jsonl_chunks.extend(chunks)

        # JSONL一括出力
        if opts.output_format == "jsonl" and all_jsonl_chunks:
            jsonl_file = self.output_dir / f"{excel_path.stem}_rag.jsonl"
            await asyncio.to_thread(JsonlExporter.export, all_jsonl_chunks, jsonl_file)

        return sheet_results, structured_data
