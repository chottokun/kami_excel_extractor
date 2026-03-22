import json
import logging
import base64
import asyncio
<<<<<<< HEAD
=======
import os
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
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
        
        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)
        self._image_cache = {}
        
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

    def _encode_image_to_base64_url(self, image_path: Path):
<<<<<<< HEAD
        abs_path = str(image_path.absolute())
        if abs_path in self._image_cache:
            return self._image_cache[abs_path]

        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("utf-8")
            result = f"data:image/png;base64,{encoded}"
            self._image_cache[abs_path] = result
            return result

    def extract_structured_data(self, excel_path: str, model: str = "gemini/gemini-1.5-flash", system_prompt: str = None, include_visual_summaries: bool = False):
        """
        Excelを解析して構造化データを取得する
        """
        return asyncio.run(self.aextract_structured_data(excel_path, model=model, system_prompt=system_prompt, include_visual_summaries=include_visual_summaries))

    async def aextract_structured_data(self, excel_path: str, model: str = "gemini/gemini-1.5-flash", system_prompt: str = None, include_visual_summaries: bool = False):
        """
        Excelを解析して構造化データを取得する（非同期版）
        """
        import os
=======
        cache_key = str(image_path)
        if cache_key in self._image_cache:
            return self._image_cache[cache_key]
            
        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("utf-8")
            result = f"data:image/png;base64,{encoded}"
            self._image_cache[cache_key] = result
            return result

    def extract_structured_data(self, excel_path: str, model: str = None, system_prompt: str = None, include_visual_summaries: bool = False):
        """Excelを解析して構造化データを取得する (同期版ラッパー)"""
        return asyncio.run(self.aextract_structured_data(
            excel_path, model=model, system_prompt=system_prompt, include_visual_summaries=include_visual_summaries
        ))

    async def aextract_structured_data(self, excel_path: str, model: str = None, system_prompt: str = None, include_visual_summaries: bool = False):
        """Excelを解析して構造化データを取得する (非同期版)"""
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
        import re
        import yaml
        
        model = model or self.default_model
        if not any(model.startswith(p) for p in ["gemini/", "openai/", "azure/"]):
             model = f"gemini/{model}"

        excel_path = Path(excel_path)
        logger.info(f"Extracting structured data from {excel_path.name} using {model}...")
        
        png_path = self.converter.convert(excel_path)
        raw_data = self.extractor.extract(excel_path)
        image_url = self._encode_image_to_base64_url(png_path)
        
        rpm_limit = int(os.getenv("GEMINI_RPM_LIMIT", "15"))
        semaphore = asyncio.Semaphore(rpm_limit) if rpm_limit > 0 else None
        
        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。提供されたHTMLテーブルのデータを統合し、意味論的に整理された構造化データをYAML形式で出力してください。出力は必ず ```yaml と ``` で囲んだブロック内のみとしてください。"

        sheets_data = raw_data.get("sheets", {})
        structured_sheets = {}
<<<<<<< HEAD

        async def process_sheet(sheet_name, sheet_content):
            logger.info(f"Processing sheet: {sheet_name}")
            
            messages = [
                {"role": "system", "content": system_prompt},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"対象シート: {sheet_name}\nデータソース:\n{sheet_content.get('html', '')}"},
                        {"type": "image_url", "image_url": {"url": image_url}},
                        {"type": "text", "text": "提供されたExcelシートのデータを構造化データとしてYAMLで出力してください。Markdownのコードブロック(```yaml)を含めてください。"}
                    ]
                }
            ]

            async def _call_llm():
                if semaphore:
                    async with semaphore:
                        return await litellm.acompletion(
                            model=model,
                            messages=messages,
                            api_key=self.api_key,
                            base_url=self.base_url
                        )
                else:
                    return await litellm.acompletion(
=======

        async def process_sheet(sheet_name, sheet_content):
            async with (semaphore if semaphore else asyncio.Lock()):
                logger.info(f"Processing sheet: {sheet_name}")
                
                messages = [
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": f"対象シート: {sheet_name}\nデータソース:\n{sheet_content.get('html', '')}"},
                            {"type": "image_url", "image_url": {"url": image_url}},
                            {"type": "text", "text": "提供されたExcelシートのデータを構造化データとしてYAMLで出力してください。Markdownのコードブロック(```yaml)を含めてください。"}
                        ]
                    }
                ]

                yaml_str = ""
                try:
                    response = await litellm.acompletion(
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
                        model=model,
                        messages=messages,
                        api_key=self.api_key,
                        base_url=self.base_url
                    )
<<<<<<< HEAD

            yaml_str = ""
            try:
                response = await _call_llm()
                raw_content = response.choices[0].message.content
                
                yaml_str = raw_content
                yaml_match = re.search(r'```(?:yaml|yml)\n(.*?)\n```', raw_content, re.DOTALL)
                if yaml_match:
                    yaml_str = yaml_match.group(1)
                elif raw_content.startswith('```'):
                    yaml_str = re.sub(r'^```[\w]*\n|\n```$', '', raw_content)
                
                sheet_data = yaml.safe_load(yaml_str) or {}
                if not isinstance(sheet_data, dict):
                    sheet_data = {"data": sheet_data}
                    
                if "sheets" in sheet_data and sheet_name in sheet_data["sheets"]:
                    res = sheet_data["sheets"][sheet_name]
                else:
                    res = sheet_data
                    
                res["_raw_yaml"] = yaml_str
                return sheet_name, res
                    
            except Exception as e:
                logger.error(f"Structured extraction (YAML) failed for sheet {sheet_name}: {e}")
                return sheet_name, {"error": str(e), "_raw_yaml": yaml_str}

        sheet_tasks = [process_sheet(name, content) for name, content in sheets_data.items()]
        sheet_results = await asyncio.gather(*sheet_tasks)
        for name, res in sheet_results:
=======
                    
                    raw_content = response.choices[0].message.content
                    yaml_match = re.search(r'```(?:yaml|yml)\n(.*?)\n```', raw_content, re.DOTALL)
                    yaml_str = yaml_match.group(1) if yaml_match else raw_content
                    
                    sheet_data = yaml.safe_load(yaml_str) or {}
                    if not isinstance(sheet_data, dict):
                        sheet_data = {"data": sheet_data}
                        
                    if "sheets" in sheet_data and sheet_name in sheet_data["sheets"]:
                        result = sheet_data["sheets"][sheet_name]
                    else:
                        result = sheet_data
                    
                    result["_raw_yaml"] = yaml_str
                    return sheet_name, result
                except Exception as e:
                    logger.error(f"Structured extraction (YAML) failed for sheet {sheet_name}: {e}")
                    return sheet_name, {"error": str(e), "_raw_yaml": yaml_str}

        tasks = [process_sheet(name, content) for name, content in sheets_data.items()]
        results = await asyncio.gather(*tasks)
        for name, res in results:
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
            structured_sheets[name] = res

        structured_data = {"sheets": structured_sheets}
        
        if include_visual_summaries:
<<<<<<< HEAD
            all_media = []
            logger.info(f"raw_data sheets: {list(sheets_data.keys())}")

            async def process_media_item(media_item, sheet_name):
                media_path = self.output_dir / "media" / media_item["filename"]
                if media_path.exists():
                    logger.info(f"Found media file: {media_path}. Generating summary...")

                    async def _call_vsum():
                        if semaphore:
                            async with semaphore:
                                return await self.aget_visual_summary(media_path, model=model)
                        else:
                            return await self.aget_visual_summary(media_path, model=model)

                    summary = await _call_vsum()
                    media_item["visual_summary"] = summary
                    return media_item, sheet_name
                else:
                    logger.error(f"Media file NOT FOUND: {media_path}")
                    return None, sheet_name

            media_tasks = []
            for sheet_name, sheet_data in sheets_data.items():
                media_list = sheet_data.get("media", [])
                logger.info(f"Sheet '{sheet_name}' has {len(media_list)} media items.")
                for media_item in media_list:
                    media_tasks.append(process_media_item(media_item, sheet_name))

            media_results = await asyncio.gather(*media_tasks)

            sheet_to_media = {}
            for media_item, sheet_name in media_results:
                if media_item:
                    all_media.append(media_item)
                    if sheet_name not in sheet_to_media:
                        sheet_to_media[sheet_name] = []
                    sheet_to_media[sheet_name].append(media_item)

            for sheet_name, items in sheet_to_media.items():
                if sheet_name in structured_sheets:
                    structured_sheets[sheet_name]["media"] = items
            
            if all_media:
                structured_data["media"] = all_media
                logger.info(f"Successfully integrated {len(all_media)} media summaries.")
            else:
                logger.warning("No media items were collected for visual summaries.")
                
        return structured_data

    def get_visual_summary(self, image_path: Path, model: str = "gemini/gemini-1.5-flash") -> str:
        """画像を個別に解析して要約文を生成する"""
        return asyncio.run(self.aget_visual_summary(image_path, model=model))

    async def aget_visual_summary(self, image_path: Path, model: str = "gemini/gemini-1.5-flash") -> str:
        """画像を個別に解析して要約文を生成する(非同期版)"""
        logger.info(f"Generating visual summary for {image_path.name}...")
=======
            async def process_media(media_item, sheet_name):
                media_path = self.output_dir / "media" / media_item["filename"]
                if media_path.exists():
                    async with (semaphore if semaphore else asyncio.Lock()):
                        summary = await self.aget_visual_summary(media_path, model=model)
                        media_item["visual_summary"] = summary
                        return media_item
                return None

            all_media_tasks = []
            for sheet_name, sheet_data in sheets_data.items():
                for media_item in sheet_data.get("media", []):
                    all_media_tasks.append(process_media(media_item, sheet_name))
            
            if all_media_tasks:
                media_results = await asyncio.gather(*all_media_tasks)
                all_media = [m for m in media_results if m]
                structured_data["media"] = all_media
                
                for m in all_media:
                    sheet_name_part = m["filename"].split("_img_")[0] if "_img_" in m["filename"] else m["filename"].split("_")[0]
                    if sheet_name_part in structured_sheets:
                        if "media" not in structured_sheets[sheet_name_part]:
                            structured_sheets[sheet_name_part]["media"] = []
                        structured_sheets[sheet_name_part]["media"].append(m)
                
        # 最終出力をJSONシリアライズ可能な形式に変換
        return self._make_json_serializable(structured_data)

    def get_visual_summary(self, image_path: Path, model: str = None) -> str:
        """画像を個別に解析して要約文を生成する (同期版ラッパー)"""
        return asyncio.run(self.aget_visual_summary(image_path, model=model))

    async def aget_visual_summary(self, image_path: Path, model: str = None) -> str:
        """画像を個別に解析して要約文を生成する (非同期版)"""
        model = model or self.default_model
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
        image_url = self._encode_image_to_base64_url(image_path)
        prompt = "この画像・グラフの内容を詳細に説明してください。回答の冒頭に [画像概要] と付けて出力してください。"
        
        messages = [{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": image_url}}]}]
        
        try:
<<<<<<< HEAD
            response = await litellm.acompletion(
                model=model,
                messages=messages,
                api_key=self.api_key,
                base_url=self.base_url
            )
=======
            response = await litellm.acompletion(model=model, messages=messages, api_key=self.api_key, base_url=self.base_url)
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
            return response.choices[0].message.content or ""
        except Exception as e:
            logger.error(f"Failed to generate visual summary: {e}")
            return "[画像概要] 解析に失敗しました。"

    def extract_rag_chunks(self, excel_path: str, model: str = None, list_format: str = "kv"):
        """Excelを解析してRAG用のチャンクを生成する (同期版ラッパー)"""
        return asyncio.run(self.aextract_rag_chunks(excel_path, model=model, list_format=list_format))

    async def aextract_rag_chunks(self, excel_path: str, model: str = None, list_format: str = "kv"):
        """Excelを解析してRAG用のチャンクを生成する (非同期版)"""
        excel_path = Path(excel_path)
        raw_data = self.extractor.extract(excel_path)
        has_simple_table = any(s.get("is_simple") for s in raw_data.get("sheets", {}).values())
        
        structured_data = await self.aextract_structured_data(excel_path, model=model, include_visual_summaries=True)

        if has_simple_table:
            for sheet_name, sheet_info in raw_data.get("sheets", {}).items():
                if sheet_info.get("is_simple"):
                    if sheet_name in structured_data["sheets"]:
                        logger.info(f"Replacing VLM output for simple table sheet '{sheet_name}'")
                        structured_data["sheets"][sheet_name] = sheet_info["structured_data"]
                    else:
                        structured_data["sheets"][sheet_name] = sheet_info["structured_data"]

        original_format = self.rag_converter.list_format
        self.rag_converter.list_format = list_format
        
        sheet_results = {}
        try:
            for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
                yaml_source = ""
                # sheet_data は _make_json_serializable 通過後も dict のはず
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
