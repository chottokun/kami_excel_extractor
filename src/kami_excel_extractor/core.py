import logging
import base64
import asyncio
import os
import re
import json
import yaml
from typing import Optional
from pathlib import Path
from datetime import date, datetime
import litellm
from .extractor import MetadataExtractor
from .converter import ExcelConverter
from .rag_converter import JsonToMarkdownConverter, RagChunker
from .document_generator import DocumentGenerator
from .schema import SheetData, ExtractionOptions, RagOptions

logger = logging.getLogger(__name__)

class KamiExcelExtractor:
    """Excelから構造化JSONを抽出するメインクラス（OpenAI / Gemini / Azure対応）"""

    _RE_YAML_BLOCK = re.compile(r'```(?:yaml|yml)\n(.*?)\n```', re.DOTALL)
    _RE_JSON_BLOCK = re.compile(r'```json\n(.*?)\n```', re.DOTALL)

    def __init__(self, api_key: str = None, output_dir: str = "output", base_url: str = None, timeout: float = 600.0):
        """
        Args:
            api_key: LLMのAPIキー
            output_dir: 結果の出力先
            base_url: プロキシ等を使用する場合のカスタムURL
            timeout: 推論のタイムアウト（秒）
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.timeout = timeout if timeout != 600.0 else float(os.getenv("LLM_TIMEOUT") or timeout)
        
        # 汎用設定を優先
        self.base_url = base_url or os.getenv("LLM_BASE_URL") or os.getenv("OLLAMA_BASE_URL")
        
        # RPM制限の初期化
        rpm_str = os.getenv("LLM_RPM_LIMIT") or os.getenv("GEMINI_RPM_LIMIT") or "15"
        self.litellm_rpm_limit = int(rpm_str)
        self._image_cache = {} # PR #16: Image cache
        self._visual_summary_cache = {} # Visual summary cache

        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)

        
        # モデル名の取得 (汎用環境変数 LLM_MODEL を最優先)
        env_model = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"
        
        # プレフィックスが含まれていない場合、かつ Gemini でない可能性がある場合は注意が必要だが、
        # 基本的には指定通りに扱う (litellmに任せる)
        self.default_model = env_model

        if api_key:
            self.api_key = api_key.strip("'\" ")
        else:
            self.api_key = os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")

    def _make_json_serializable(self, data):
        """
        辞書内の date/datetime オブジェクトを文字列に変換する (再帰的)
        """
        if isinstance(data, dict):
            return {k: self._make_json_serializable(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self._make_json_serializable(i) for i in data]
        elif isinstance(data, (date, datetime)):
            return data.isoformat()
        return data

    async def _encode_image_to_base64_url(self, image_path: Path):
        cache_key = str(image_path)
        if cache_key in self._image_cache:
            return self._image_cache[cache_key]
            
        def _read_and_encode():
            with open(image_path, "rb") as image_file:
                encoded = base64.b64encode(image_file.read()).decode("utf-8")
                return f"data:image/png;base64,{encoded}"

        result = await asyncio.to_thread(_read_and_encode)
        self._image_cache[cache_key] = result
        return result

    def _resolve_model(self, model: Optional[str] = None) -> str:
        """モデル名のデフォルト値とプレフィックスを解決する"""
        model = model or self.default_model
        # プレフィックスの自動付加は廃止 (litellmの推論や明示的な指定に任せる)
        return model

    def _get_semaphore(self) -> Optional[asyncio.Semaphore]:
        """RPM制限に基づいたセマフォを取得する"""
        return asyncio.Semaphore(self.litellm_rpm_limit) if self.litellm_rpm_limit > 0 else None

    def _build_sheet_messages(self, system_prompt: str, sheet_name: str, html_content: str, image_url: Optional[str] = None) -> list:
        """LLMへのメッセージリストを構築する"""
        content = [{"type": "text", "text": f"対象シート: {sheet_name}\nデータソース:\n{html_content}"}]
        
        if image_url:
            content.append({"type": "image_url", "image_url": {"url": image_url}})
            
        content.append({"type": "text", "text": "提供されたExcelシートのデータを解析し、構造化されたJSONオブジェクトとして出力してください。Markdownのコードブロック(```json)を含めてください。"})
        
        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content}
        ]

    def _parse_llm_response(self, content: str, sheet_name: str) -> dict:
        """LLMのレスポンスからJSONまたはYAMLを抽出してパースし、Pydanticで検証する"""
        # JSONブロックの抽出を優先
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

        # JSONで見つからない、またはパース失敗した場合はYAMLを試行
        if data is None:
            yaml_match = self._RE_YAML_BLOCK.search(content)
            yaml_str = yaml_match.group(1) if yaml_match else content
            try:
                data = yaml.safe_load(yaml_str)
                raw_str = yaml_str
            except Exception as e:
                logger.error(f"Failed to parse YAML for sheet {sheet_name}: {e}")
                return {"error": str(e), "_raw_data": content}

        try:
            # データの整形
            if not isinstance(data, dict):
                data = {"data": data}
            
            # 階層構造の正規化 (sheets[sheet_name] 形式の場合)
            if "sheets" in data and sheet_name in data["sheets"]:
                data = data["sheets"][sheet_name]

            # Pydanticによるバリデーション
            validated = SheetData(**data)
            result = validated.model_dump()
            result["_raw_data"] = raw_str
            return result
        except Exception as e:
            logger.error(f"Validation failed for sheet {sheet_name}: {e}")
            return {"error": f"Validation failed: {str(e)}", "_raw_data": raw_str}

    async def _aextract_single_sheet(self, sheet_name: str, sheet_content: dict, options: ExtractionOptions, image_url: str, semaphore: Optional[asyncio.Semaphore]) -> tuple:
        """単一のシートを解析して構造化データを取得する（リトライロジック付き）"""
        if sheet_content.get("is_simple"):
            logger.info(f"Using simple table extraction for sheet: {sheet_name}")
            result = sheet_content.get("structured_data", [])
            if not isinstance(result, dict):
                result = {"data": result}
            result["_raw_data"] = ""
            return sheet_name, result

        async with (semaphore if semaphore else asyncio.Lock()):
            logger.info(f"Processing sheet via LLM: {sheet_name}")
            actual_image_url = image_url if options.use_visual_context else None
            messages = self._build_sheet_messages(options.system_prompt, sheet_name, sheet_content.get('html', ''), actual_image_url)

            # リトライループ
            max_retries = 1
            last_error = None
            
            for attempt in range(max_retries + 1):
                if attempt > 0:
                    logger.warning(f"Retry attempt {attempt} for sheet: {sheet_name} due to error: {last_error}")
                    # エラーフィードバックを含めたメッセージを構築
                    messages.append({"role": "assistant", "content": f"エラー内容: {last_error}\n恐れ入りますが、上記のパースエラーを修正した正しいJSON形式で再度出力してください。"})

                try:
                    # JSON出力の強制 (モデルがサポートしている場合)
                    response_format = {"type": "json_object"} if "ollama" in options.model or "gpt" in options.model or "gemini" in options.model else None
                    
                    response = await litellm.acompletion(
                        model=options.model,
                        messages=messages,
                        api_key=self.api_key,
                        base_url=self.base_url,
                        timeout=self.timeout,
                        response_format=response_format
                    )
                    content = response.choices[0].message.content
                    result = self._parse_llm_response(content, sheet_name)
                    
                    if "error" not in result:
                        return sheet_name, result
                    
                    last_error = result["error"]
                except Exception as e:
                    logger.error(f"Extraction failed for sheet {sheet_name} (Attempt {attempt}): {e}")
                    last_error = str(e)
                
            # リトライ後も失敗した場合
            return sheet_name, {"error": f"Failed after {max_retries} retries. Last error: {last_error}", "_raw_data": ""}

    async def _aprocess_media_summary(self, media_item: dict, model: str, semaphore: Optional[asyncio.Semaphore]) -> Optional[dict]:
        """単一のメディアアイテムの要約文を生成する"""
        media_path = self.output_dir / "media" / media_item["filename"]
        if not media_path.exists():
            return None

        async with (semaphore if semaphore else asyncio.Lock()):
            summary = await self.aget_visual_summary(media_path, model=model)
            media_item["visual_summary"] = summary
            return media_item

    def _attach_media_to_sheets(self, all_media: list, structured_sheets: dict) -> None:
        """抽出されたメディア情報を各シートのデータに紐付ける"""
        for m in all_media:
            # ファイル名からシート名を推測 (Sheet1_img_A1_0.png -> Sheet1)
            filename = m.get("filename", "")
            sheet_name_part = filename.split("_img_")[0] if "_img_" in filename else filename.split("_")[0]

            if sheet_name_part in structured_sheets:
                sheet_struct = structured_sheets[sheet_name_part]
                if isinstance(sheet_struct, dict):
                    if "media" not in sheet_struct:
                        sheet_struct["media"] = []
                    sheet_struct["media"].append(m)

    def extract_structured_data(self, excel_path: str, options: Optional[ExtractionOptions] = None):
        """Excelを解析して構造化データを取得する (同期版ラッパー)"""
        return asyncio.run(self.aextract_structured_data(excel_path, options=options))

    async def aextract_structured_data(self, excel_path: str, options: Optional[ExtractionOptions] = None):
        """Excelを解析して構造化データを取得する (非同期版)"""
        options = options or ExtractionOptions()
        model = self._resolve_model(options.model)
        options.model = model # 解決済みモデルで更新

        semaphore = self._get_semaphore()
        excel_path = Path(excel_path)
        logger.info(f"Extracting structured data from {excel_path.name} using {model}...")

        # 変換・抽出・画像URL化
        raw_data = self.extractor.extract(excel_path)
        
        image_url = None
        if options.include_visual_summaries or options.use_visual_context:
            logger.info("Converting Excel to image for visual context/summaries...")
            try:
                png_path = self.converter.convert(excel_path)
                image_url = await self._encode_image_to_base64_url(png_path)
            except Exception as e:
                logger.warning(f"Excel-to-image conversion failed (skipping visual context): {e}")
        
        if not options.system_prompt:
            options.system_prompt = "あなたはExcel構造化の専門家です。提供されたHTMLテーブルのデータを統合し、意味論的に整理された構造化データをJSON形式で出力してください。出力は必ず ```json と ``` で囲んだブロック内のみとしてください。"

        # 各シートの解析タスクの実行
        sheets_data = raw_data.get("sheets", {})
        tasks = [
            self._aextract_single_sheet(name, content, options, image_url, semaphore)
            for name, content in sheets_data.items()
        ]
        results = await asyncio.gather(*tasks)
        structured_sheets = {name: res for name, res in results}
        structured_data = {"sheets": structured_sheets}

        # ビジュアルサマリー（画像要約）の生成と紐付け
        if options.include_visual_summaries:
            media_tasks = []
            for sheet_name, sheet_info in sheets_data.items():
                for media_item in sheet_info.get("media", []):
                    media_tasks.append(self._aprocess_media_summary(media_item, options.model, semaphore))
            
            if media_tasks:
                media_results = await asyncio.gather(*media_tasks)
                all_media = [m for m in media_results if m]
                structured_data["media"] = all_media
                self._attach_media_to_sheets(all_media, structured_sheets)
                
        # 最終出力をJSONシリアライズ可能な形式に変換
        return self._make_json_serializable(structured_data)

    def get_visual_summary(self, image_path: Path, model: str = None) -> str:
        """画像を個別に解析して要約文を生成する (同期版ラッパー)"""
        return asyncio.run(self.aget_visual_summary(image_path, model=model))

    async def aget_visual_summary(self, image_path: Path, model: str = None) -> str:
        """画像を個別に解析して要約文を生成する (非同期版)"""
        model = model or self.default_model
        image_url = await self._encode_image_to_base64_url(image_path)
        
        cache_key = (model, image_url)
        if cache_key in self._visual_summary_cache:
            return self._visual_summary_cache[cache_key]
            
        prompt = "この画像・グラフの内容を詳細に説明してください。回答の冒頭に [画像概要] と付けて出力してください。"
        
        messages = [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": image_url}}]}]
        
        try:
            response = await litellm.acompletion(
                model=model,
                messages=messages,
                api_key=self.api_key,
                base_url=self.base_url,
                timeout=self.timeout
            )
            
            content = response.choices[0].message.content or ""
            if content:
                self._visual_summary_cache[cache_key] = content
            return content
        except Exception as e:
            logger.error(f"Failed to generate visual summary: {e}")
            return "[画像概要] 解析に失敗しました。"

    def extract_rag_chunks(self, excel_path: str, options: Optional[RagOptions] = None):
        """Excelを解析してRAG用のチャンクを生成する (同期版ラッパー)"""
        return asyncio.run(self.aextract_rag_chunks(excel_path, options=options))

    async def aextract_rag_chunks(self, excel_path: str, options: Optional[RagOptions] = None):
        """Excelを解析してRAG用のチャンクを生成する (非同期版)"""
        options = options or RagOptions()
        excel_path = Path(excel_path)
        
        # RAG用なのでビジュアルサマリーは強制的にTrue
        options.include_visual_summaries = True

        structured_data = await self.aextract_structured_data(excel_path, options=options)

        original_format = self.rag_converter.list_format
        self.rag_converter.list_format = options.list_format
        
        sheet_results = {}
        try:
            for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
                raw_data_source = ""
                if isinstance(sheet_data, dict):
                    raw_data_source = sheet_data.pop("_raw_data", "")
                
                single_sheet_data = {"sheets": {sheet_name: sheet_data}}
                if "media" in structured_data:
                    single_sheet_data["media"] = [m for m in structured_data["media"] if m.get("filename", "").startswith(sheet_name)]
                
                markdown_text = self.rag_converter.convert(single_sheet_data)
                
                metadata = {"source_file": excel_path.name, "sheet_name": sheet_name, "extraction_model": options.model or self.default_model}
                chunker = RagChunker(metadata=metadata)
                chunks = chunker.chunk(markdown_text, f"{excel_path.name} - {sheet_name}")
                
                sheet_results[sheet_name] = {
                    "chunks": chunks,
                    "markdown": markdown_text,
                    "structured": single_sheet_data,
                    "raw_data": raw_data_source
                }
        finally:
            self.rag_converter.list_format = original_format
            
        return sheet_results, structured_data
