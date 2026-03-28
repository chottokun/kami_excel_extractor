import json
import logging
import base64
import asyncio
import os
import re
import yaml
from typing import Optional, Any
from pathlib import Path
from datetime import date, datetime
import litellm
from .extractor import MetadataExtractor
from .converter import ExcelConverter
from .rag_converter import JsonToMarkdownConverter, RagChunker
from .document_generator import DocumentGenerator

logger = logging.getLogger(__name__)

class KamiExcelExtractor:
    """Excelから構造化JSONを抽出するメインクラス（OpenAI / Gemini / Azure対応）"""

    def __init__(self, api_key: str = None, output_dir: str = "output", base_url: str = None):
        """
        Args:
            api_key: LLMのAPIキー
            output_dir: 結果の出力先
            base_url: プロキシ等を使用する場合のカスタムURL
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.base_url = base_url
        self._image_cache = {} # PR #16: Image cache
        self._visual_summary_cache = {} # Visual summary cache

        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)

        
        # モデル名の取得 (環境変数を優先)
        env_model = os.getenv("GEMINI_MODEL", "gemini-1.5-flash")
        self.default_model = env_model if env_model.startswith(("gemini/", "openai/")) else f"gemini/{env_model}"

        if api_key:
            self.api_key = api_key.strip("'\" ")
        else:
            self.api_key = os.getenv("GEMINI_API_KEY")

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
        if not any(model.startswith(p) for p in ["gemini/", "openai/", "azure/"]):
            model = f"gemini/{model}"
        return model

    def _get_semaphore(self) -> Optional[asyncio.Semaphore]:
        """RPM制限に基づいたセマフォを取得する"""
        rpm_limit = int(os.getenv("GEMINI_RPM_LIMIT", "15"))
        return asyncio.Semaphore(rpm_limit) if rpm_limit > 0 else None

    def _build_sheet_messages(self, system_prompt: str, sheet_name: str, html_content: str, image_url: str) -> list:
        """LLMへのメッセージリストを構築する"""
        return [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"対象シート: {sheet_name}\nデータソース:\n{html_content}"},
                    {"type": "image_url", "image_url": {"url": image_url}},
                    {"type": "text", "text": "提供されたExcelシートのデータを構造化データとしてYAMLで出力してください。Markdownのコードブロック(```yaml)を含めてください。"}
                ]
            }
        ]

    def _parse_yaml_response(self, content: str, sheet_name: str) -> dict:
        """LLMのレスポンスからYAMLを抽出してパースする"""
        yaml_match = re.search(r'```(?:yaml|yml)\n(.*?)\n```', content, re.DOTALL)
        yaml_str = yaml_match.group(1) if yaml_match else content

        try:
            sheet_data = yaml.safe_load(yaml_str) or {}
            if not isinstance(sheet_data, dict):
                sheet_data = {"data": sheet_data}

            if "sheets" in sheet_data and sheet_name in sheet_data["sheets"]:
                result = sheet_data["sheets"][sheet_name]
            else:
                result = sheet_data

            result["_raw_yaml"] = yaml_str
            return result
        except Exception as e:
            logger.error(f"Failed to parse YAML for sheet {sheet_name}: {e}")
            return {"error": str(e), "_raw_yaml": yaml_str}

    async def _aextract_single_sheet(self, sheet_name: str, sheet_content: dict, model: str, system_prompt: str, image_url: str, semaphore: Optional[asyncio.Semaphore]) -> tuple:
        """単一のシートを解析して構造化データを取得する"""
        if sheet_content.get("is_simple"):
            logger.info(f"Using simple table extraction for sheet: {sheet_name}")
            # simple table の場合は _raw_yaml は空か、あるいは生成したものを入れる
            result = sheet_content.get("structured_data", [])
            if not isinstance(result, dict):
                result = {"data": result}
            result["_raw_yaml"] = ""
            return sheet_name, result

        async with (semaphore if semaphore else asyncio.Lock()):
            logger.info(f"Processing sheet via LLM: {sheet_name}")
            messages = self._build_sheet_messages(system_prompt, sheet_name, sheet_content.get('html', ''), image_url)

            try:
                response = await litellm.acompletion(
                    model=model,
                    messages=messages,
                    api_key=self.api_key,
                    base_url=self.base_url
                )
                result = self._parse_yaml_response(response.choices[0].message.content, sheet_name)
                return sheet_name, result
            except Exception as e:
                logger.error(f"Structured extraction (YAML) failed for sheet {sheet_name}: {e}")
                return sheet_name, {"error": str(e), "_raw_yaml": ""}

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

    def extract_structured_data(self, excel_path: str, model: str = None, system_prompt: str = None, include_visual_summaries: bool = False):
        """Excelを解析して構造化データを取得する (同期版ラッパー)"""
        return asyncio.run(self.aextract_structured_data(
            excel_path, model=model, system_prompt=system_prompt, include_visual_summaries=include_visual_summaries
        ))

    async def aextract_structured_data(self, excel_path: str, model: str = None, system_prompt: str = None, include_visual_summaries: bool = False):
        """Excelを解析して構造化データを取得する (非同期版)"""
        model = self._resolve_model(model)
        semaphore = self._get_semaphore()
        excel_path = Path(excel_path)
        logger.info(f"Extracting structured data from {excel_path.name} using {model}...")

        # 変換・抽出・画像URL化
        png_path = self.converter.convert(excel_path)
        raw_data = self.extractor.extract(excel_path)
        image_url = await self._encode_image_to_base64_url(png_path)
        
        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。提供されたHTMLテーブルのデータを統合し、意味論的に整理された構造化データをYAML形式で出力してください。出力は必ず ```yaml と ``` で囲んだブロック内のみとしてください。"

        # 各シートの解析タスクの実行
        sheets_data = raw_data.get("sheets", {})
        tasks = [
            self._aextract_single_sheet(name, content, model, system_prompt, image_url, semaphore)
            for name, content in sheets_data.items()
        ]
        results = await asyncio.gather(*tasks)
        structured_sheets = {name: res for name, res in results}
        structured_data = {"sheets": structured_sheets}

        # ビジュアルサマリー（画像要約）の生成と紐付け
        if include_visual_summaries:
            media_tasks = []
            for sheet_name, sheet_info in sheets_data.items():
                for media_item in sheet_info.get("media", []):
                    media_tasks.append(self._aprocess_media_summary(media_item, model, semaphore))
            
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
            response = await litellm.acompletion(model=model, messages=messages, api_key=self.api_key, base_url=self.base_url)
            
            content = response.choices[0].message.content or ""
            if content:
                self._visual_summary_cache[cache_key] = content
            return content
        except Exception as e:
            logger.error(f"Failed to generate visual summary: {e}")
            return "[画像概要] 解析に失敗しました。"

    def extract_rag_chunks(self, excel_path: str, model: str = None, list_format: str = "kv"):
        """Excelを解析してRAG用のチャンクを生成する (同期版ラッパー)"""
        return asyncio.run(self.aextract_rag_chunks(excel_path, model=model, list_format=list_format))

    async def aextract_rag_chunks(self, excel_path: str, model: str = None, list_format: str = "kv"):
        """Excelを解析してRAG用のチャンクを生成する (非同期版)"""
        excel_path = Path(excel_path)
        
        structured_data = await self.aextract_structured_data(excel_path, model=model, include_visual_summaries=True)

        original_format = self.rag_converter.list_format
        self.rag_converter.list_format = list_format
        
        sheet_results = {}
        try:
            for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
                yaml_source = ""
                if isinstance(sheet_data, dict):
                    yaml_source = sheet_data.pop("_raw_yaml", "")
                
                single_sheet_data = {"sheets": {sheet_name: sheet_data}}
                if "media" in structured_data:
                    single_sheet_data["media"] = [m for m in structured_data["media"] if m.get("filename", "").startswith(sheet_name)]
                
                markdown_text = self.rag_converter.convert(single_sheet_data)
                
                metadata = {"source_file": excel_path.name, "sheet_name": sheet_name, "extraction_model": model or self.default_model}
                chunker = RagChunker(metadata=metadata)
                chunks = chunker.chunk(markdown_text, f"{excel_path.name} - {sheet_name}")
                
                sheet_results[sheet_name] = {
                    "chunks": chunks,
                    "markdown": markdown_text,
                    "structured": single_sheet_data,
                    "yaml": yaml_source
                }
        finally:
            self.rag_converter.list_format = original_format
            
        return sheet_results, structured_data
