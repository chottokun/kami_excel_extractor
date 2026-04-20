"""
Kami Excel Extractor のコアオーケストレーションエンジン。
マルチモーダルLLM、視覚解析、およびロジック解析を統合し、構造化データを生成する。
"""

import asyncio
import aiofiles
import base64
import json
import logging
import os
import re
import yaml
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import litellm
from pydantic import ValidationError

from .converter import ExcelConverter
from .document_generator import DocumentGenerator
from .extractor import MetadataExtractor
from .rag_converter import JsonToMarkdownConverter, RagChunker
from .schema import ExtractionOptions, ExtractionResult, FullExtraction, RagOptions, SheetData

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
        timeout: int = 600,
        litellm_rpm_limit: int = 0
    ):
        """
        KamiExcelExtractor を初期化する。
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.base_url = base_url
        self.timeout = timeout
        self.litellm_rpm_limit = litellm_rpm_limit or int(os.getenv("LLM_RPM_LIMIT", 0))
        
        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)

        self._image_cache = {}
        self._image_locks = {}
        self._visual_summary_cache = {}
        
        self.default_model = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"
        self.api_key = api_key.strip("'\" ") if api_key else (os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY"))

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
        """画像をBase64エンコードし、データURL形式で返す。"""
        cache_key = str(image_path)
        if cache_key in self._image_cache:
            return self._image_cache[cache_key]
            
        # 同一ファイルへの同時エンコードを防ぐためのロック取得
        lock = self._image_locks.setdefault(cache_key, asyncio.Lock())
        async with lock:
            if cache_key in self._image_cache:
                return self._image_cache[cache_key]

            async with aiofiles.open(image_path, "rb") as f:
                content = await f.read()

            def _encode(data: bytes) -> str:
                encoded = base64.b64encode(data).decode("utf-8")
                return f"data:image/png;base64,{encoded}"

            result = await asyncio.to_thread(_encode, content)
            self._image_cache[cache_key] = result
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
                    delay = initial_delay * (2 ** attempt)
                    logger.warning(f"API error ({status_code}). Retrying in {delay:.1f}s (Attempt {attempt + 1}/{max_retries})...")
                    await asyncio.sleep(delay)
                else:
                    break
        raise last_exception

    def _get_semaphore(self) -> asyncio.Semaphore:
        """RPM制限用のセマフォを取得する。制限なし(0)の場合は十分に大きな値を設定。"""
        limit = self.litellm_rpm_limit if self.litellm_rpm_limit > 0 else 1000
        return asyncio.Semaphore(limit)

    def _build_sheet_messages(self, system_prompt: str, sheet_name: str, html_content: str, image_url: Optional[str] = None, include_logic: bool = False) -> List[Dict]:
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
        
        if image_url:
            content.append({"type": "image_url", "image_url": {"url": image_url}})
            
        content.append({"type": "text", "text": "解析結果を構造化されたJSONオブジェクトとして出力してください。必ず ```json ブロックを含めてください。"})
        
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
            except json.JSONDecodeError:
                pass

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
            if not isinstance(data, dict): data = {"data": data}
            if "sheets" in data and sheet_name in data["sheets"]: data = data["sheets"][sheet_name]
            # 生の応答テキストを保持 (テストおよびデバッグ用)
            data["_raw_data"] = raw_str
            # Pydanticによるスキーマ検証
            ExtractionResult(**data)
            return data
        except Exception as e:
            logger.error(f"Validation failed for {sheet_name}: {e}")
            return {"error": f"Validation failed: {str(e)}", "_raw_data": raw_str}

    async def _aprocess_chart_data(self, media_item: Dict, model: str, semaphore: Optional[asyncio.Semaphore]) -> Dict:
        """画像がグラフや図表の場合、VLMを用いてその内容を構造化データとして抽出する。"""
        if not media_item.get("filename"): return media_item
        image_path = self.output_dir / "media" / media_item["filename"]
        if not image_path.exists(): return media_item

        async with (semaphore if semaphore else asyncio.Lock()):
            image_url = await self._encode_image_to_base64_url(image_path)
            prompt = (
                "この画像がグラフや図表の場合、その軸ラベル、凡例、およびデータ値を抽出し、"
                "Markdownのテーブル形式で整理してください。回答の冒頭に [図表データ] と付けて出力してください。"
            )
            messages = [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": image_url}}]}]
            
            try:
                response = await self._awith_retry(
                    litellm.acompletion,
                    model=model, messages=messages, api_key=self.api_key, 
                    base_url=self.base_url, timeout=self.timeout
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
            f"<div class='visual-insight'>[図表データ({coord})]: {i['visual_data']}</div>"
            for i in items if "visual_data" in i
        ]
        return "\n".join(insights)

    def _inject_insight_to_html(self, html_content: str, coord: str, insight_html: str) -> str:
        """HTML内の特定の座標属性を持つ箇所にインサイトを注入する。"""
        for attr_pattern in [f'data-coord="{coord}"', f"data-coord='{coord}'"]:
            if attr_pattern in html_content:
                # セル属性の直後にデータを挿入
                return html_content.replace(attr_pattern, f"{attr_pattern} {insight_html}")
        return html_content

    async def _aprocess_media_summary(self, media_item: Dict, model: str, semaphore: Optional[asyncio.Semaphore]) -> Optional[Dict]:
        """メディアアイテムの要約文と図表データの抽出を並列実行する。"""
        if not media_item.get("filename"): return media_item
        media_path = self.output_dir / "media" / media_item["filename"]
        if not media_path.exists(): return None

        # aget_visual_summary と _aprocess_chart_data 内部でそれぞれセマフォを制御
        summary = await self.aget_visual_summary(media_path, model=model, semaphore=semaphore)
        media_item["visual_summary"] = summary
        await self._aprocess_chart_data(media_item, model, semaphore)
        return media_item

    async def _aextract_single_sheet(self, sheet_name: str, sheet_content: Dict, model: str, system_prompt: str, image_url: Optional[str], semaphore: Optional[asyncio.Semaphore], use_visual_context: bool = True, include_logic: bool = False) -> Tuple[str, Dict]:
        """単一シートの解析。シンプルテーブルの場合はLLMをバイパスする。"""
        if sheet_content.get("is_simple"):
            logger.info(f"Using simple table extraction for: {sheet_name}")
            result = sheet_content.get("structured_data", [])
            return sheet_name, {"data": result if isinstance(result, dict) else {"data": result}, "_raw_data": ""}

        async with (semaphore if semaphore else asyncio.Lock()):
            logger.info(f"Processing via LLM: {sheet_name}")
            messages = self._build_sheet_messages(system_prompt, sheet_name, sheet_content.get('html', ''), image_url if use_visual_context else None, include_logic=include_logic)
            
            max_retries = 1
            last_error = None
            for attempt in range(max_retries + 1):
                if attempt > 0:
                    messages.append({"role": "assistant", "content": f"エラー内容: {last_error}\n修正して正しいJSONで再出力してください。"})
                try:
                    fmt = {"type": "json_object"} if any(m in model for m in ["ollama", "gpt", "gemini"]) else None
                    response = await self._awith_retry(
                        litellm.acompletion,
                        model=model, messages=messages, api_key=self.api_key, 
                        base_url=self.base_url, timeout=self.timeout, response_format=fmt
                    )
                    result = self._parse_llm_response(response.choices[0].message.content, sheet_name)
                    if "error" not in result: return sheet_name, result
                    last_error = result["error"]
                except Exception as e:
                    last_error = str(e)
            return sheet_name, {"error": f"Failed after {max_retries} retries. Last error: {last_error}", "_raw_data": ""}

    async def aextract_structured_data(self, excel_path: Union[str, Path], options: Optional[ExtractionOptions] = None) -> Dict:
        """Excelを解析して構造化データを取得する (非同期エントリーポイント)。"""
        opts = options or ExtractionOptions()
        model = self._resolve_model(opts.model)
        semaphore = self._get_semaphore()
        excel_path = Path(excel_path)
        
        logger.info(f"Starting extraction for {excel_path.name} (Logic: {opts.include_logic})")

        # 1. Extractorによる基本解析
        raw_data = await asyncio.to_thread(self.extractor.extract, excel_path, include_logic=opts.include_logic)
        sheets_data = raw_data.get("sheets", {})

        # 2. 全体画像の生成
        image_url = None
        if opts.include_visual_summaries or opts.use_visual_context:
            try:
                self.converter.dpi = opts.dpi
                png_path = await asyncio.to_thread(self.converter.convert, excel_path)
                image_url = await self._encode_image_to_base64_url(png_path)
            except Exception as e:
                logger.warning(f"Excel-to-image failed: {e}")

        # 3. メディア（図表・グラフ）の個別解析
        all_media = []
        if opts.include_visual_summaries:
            # 重複メディア（同一ファイル名）を排除して解析タスクを作成
            unique_media = {}
            for s in sheets_data.values():
                for m in s.get("media", []):
                    if filename := m.get("filename"):
                        if filename not in unique_media:
                            unique_media[filename] = m

            media_tasks = [self._aprocess_media_summary(m, model, semaphore) for m in unique_media.values()]
            if media_tasks:
                media_results = await asyncio.gather(*media_tasks)
                all_media = [m for m in media_results if m]
                # 抽出された結果をメタデータに同期 (O(1) lookup で高速化)
                # (coord, filename) をキーにして、該当する全シートのメディアアイテムを保持
                lookup = {}
                for s_info in sheets_data.values():
                    for coord, mapped_list in s_info.get("media_map", {}).items():
                        for mapped_m in mapped_list:
                            if filename := mapped_m.get("filename"):
                                lookup.setdefault((coord, filename), []).append(mapped_m)

                for m in all_media:
                    filename = m.get("filename")
                    if not filename: continue

                    visual_data = m.get("visual_data")
                    visual_summary = m.get("visual_summary")

                    # lookup の全エントリを更新 (ファイル名のみをキーにする方が確実)
                    for (coord, f_name), target_list in lookup.items():
                        if f_name == filename:
                            for target in target_list:
                                if visual_data: target["visual_data"] = visual_data
                                if visual_summary: target["visual_summary"] = visual_summary

        # 4. 図表データの注入
        for sheet_name, sheet_info in sheets_data.items():
            if "media_map" in sheet_info:
                sheet_info["html"] = self._inject_visual_data_to_html(sheet_info["html"], sheet_info["media_map"])
                if "VISUAL_INSIGHT" in sheet_info["html"]: logger.warning(f"Injected visual insights for: {sheet_name}")

        sys_prompt = opts.system_prompt or "あなたはExcel構造化の専門家です。HTMLデータを統合し、意味論的なJSONを出力してください。"

        # 5. 各シートのLLM解析
        tasks = [self._aextract_single_sheet(n, c, model, sys_prompt, image_url, semaphore, use_visual_context=opts.use_visual_context, include_logic=opts.include_logic) for n, c in sheets_data.items()]
        results = await asyncio.gather(*tasks)
        structured_sheets = {name: res for name, res in results}
        
        final_data = {"sheets": structured_sheets}
        if all_media:
            final_data["media"] = all_media
            self._attach_media_to_sheets(all_media, structured_sheets)
                
        return self._make_json_serializable(final_data)

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

    async def aget_visual_summary(self, image_path: Path, model: Optional[str] = None, semaphore: Optional[asyncio.Semaphore] = None) -> str:
        """画像の視覚的要約を生成する。"""
        model = model or self.default_model
        cache_key = (model, str(image_path))
        if cache_key in self._visual_summary_cache: return self._visual_summary_cache[cache_key]

        image_url = await self._encode_image_to_base64_url(image_path)
            
        messages = [{"role": "user", "content": [{"type": "text", "text": "この画像の内容を詳細に説明してください。[画像概要] と付けて出力してください。"}, {"type": "image_url", "image_url": {"url": image_url}}]}]
        async with (semaphore if semaphore else asyncio.Lock()):
            try:
                response = await self._awith_retry(
                    litellm.acompletion,
                    model=model, messages=messages, api_key=self.api_key, 
                    base_url=self.base_url, timeout=self.timeout
                )
                content = response.choices[0].message.content or ""
                if content: self._visual_summary_cache[cache_key] = content
                return content
            except Exception as e:
                logger.error(f"Visual summary failed: {e}")
                return "[画像概要] 解析失敗。"

    def extract_structured_data(self, excel_path: Union[str, Path], options: Optional[ExtractionOptions] = None) -> Dict:
        return asyncio.run(self.aextract_structured_data(excel_path, options=options))

    def extract_rag_chunks(self, excel_path: Union[str, Path], options: Optional[RagOptions] = None) -> Tuple[Dict, Dict]:
        """Excelを解析し、RAG用のMarkdownチャンクと構造化データを同時に生成する。"""
        return asyncio.run(self.aextract_rag_chunks(excel_path, options=options))

    async def aextract_rag_chunks(self, excel_path: Union[str, Path], options: Optional[RagOptions] = None) -> Tuple[Dict, Dict]:
        opts = options or RagOptions()
        excel_path = Path(excel_path)
        extract_opts = ExtractionOptions(model=opts.model, system_prompt=opts.system_prompt, include_visual_summaries=True, use_visual_context=opts.use_visual_context, include_logic=opts.include_logic)
        
        structured_data = await self.aextract_structured_data(excel_path, options=extract_opts)
        
        sheet_results = {}
        for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
            markdown_text = self.rag_converter.convert({"sheets": {sheet_name: sheet_data}})
            chunker = RagChunker(metadata={"source": excel_path.name, "sheet": sheet_name})
            chunks = chunker.chunk(markdown_text, f"{excel_path.name}-{sheet_name}")
            sheet_results[sheet_name] = {"chunks": chunks, "markdown": markdown_text, "structured": sheet_data}
            
        return sheet_results, structured_data
