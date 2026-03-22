import json
import logging
import base64
import re
import yaml
import time
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
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
            api_key: LLMのAPIキー（Noneの場合は環境変数から読み込み）
            output_dir: 結果の出力先
            base_url: プロキシ等を使用する場合のカスタムURL
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.base_url = base_url
        
        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        self.rag_converter = JsonToMarkdownConverter()
        self.doc_generator = DocumentGenerator(self.output_dir)
        
        if api_key:
            # LiteLLMは環境変数を優先するため、明示的にセット
            self.api_key = api_key.strip("'\" ")
        else:
            self.api_key = None

    def _encode_image_to_base64_url(self, image_path: Path):
        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("utf-8")
            return f"data:image/png;base64,{encoded}"

    def extract_structured_data(self, excel_path: str, model: str = "gemini/gemini-1.5-flash", system_prompt: str = None, include_visual_summaries: bool = False):
        """
        Excelを解析して構造化データを取得する
        """
        excel_path = Path(excel_path)
        logger.info(f"Extracting structured data from {excel_path.name}...")
        
        # 1. 物理抽出とレンダリング
        png_path = self.converter.convert(excel_path)
        raw_data = self.extractor.extract(excel_path)
        image_url = self._encode_image_to_base64_url(png_path)
        
        rpm_limit = int(os.getenv("GEMINI_RPM_LIMIT", "15"))
        sleep_time = 60.0 / rpm_limit if rpm_limit > 0 else 0
        
        # 2. メッセージの構築
        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。提供されたHTMLテーブルのデータを統合し、意味論的に整理された構造化データをYAML形式で出力してください。出力は必ず ```yaml と ``` で囲んだブロック内のみとしてください。"

        structured_sheets = {}
        sheets_data = raw_data.get("sheets", {})
        
        def process_sheet(sheet_name, i):
            sheet_content = sheets_data[sheet_name]
            logger.info(f"Processing sheet: {sheet_name} ({i+1}/{len(sheets_data)})")
            
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
            # 3. LLM呼び出し
            try:
                response = litellm.completion(
                    model=model,
                    messages=messages,
                    api_key=self.api_key,
                    base_url=self.base_url
                )
                
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

        sheet_names = list(sheets_data.keys())
        if sheet_names:
            with ThreadPoolExecutor(max_workers=min(len(sheet_names), 10)) as executor:
                futures = []
                for i, sheet_name in enumerate(sheet_names):
                    futures.append(executor.submit(process_sheet, sheet_name, i))
                    if sleep_time > 0 and i < len(sheet_names) - 1:
                        time.sleep(sleep_time)
                
                for future in futures:
                    name, data = future.result()
                    structured_sheets[name] = data

        structured_data = {"sheets": structured_sheets}
        
        # 4. 画像概要の付与 (必要に応じて)
        if include_visual_summaries:
            logger.info(f"raw_data sheets: {list(sheets_data.keys())}")

            media_to_process = []
            for sheet_name, sheet_data in sheets_data.items():
                for media_item in sheet_data.get("media", []):
                    media_path = self.output_dir / "media" / media_item["filename"]
                    if media_path.exists():
                        media_to_process.append((sheet_name, media_item, media_path))
                    else:
                        logger.error(f"Media file NOT FOUND: {media_path}")

            if media_to_process:
                def process_media(item_info):
                    sheet_name, media_item, media_path = item_info
                    summary = self.get_visual_summary(media_path, model=model)
                    media_item["visual_summary"] = summary
                    return sheet_name, media_item

                with ThreadPoolExecutor(max_workers=min(len(media_to_process), 10)) as executor:
                    futures = []
                    for i, item_info in enumerate(media_to_process):
                        futures.append(executor.submit(process_media, item_info))
                        if sleep_time > 0 and i < len(media_to_process) - 1:
                            time.sleep(sleep_time)

                    all_media = []
                    sheet_media_map = {}
                    for future in futures:
                        sheet_name, media_item = future.result()
                        all_media.append(media_item)
                        if sheet_name not in sheet_media_map:
                            sheet_media_map[sheet_name] = []
                        sheet_media_map[sheet_name].append(media_item)

                    for sheet_name, media_list in sheet_media_map.items():
                        if sheet_name in structured_sheets:
                            structured_sheets[sheet_name]["media"] = media_list

                    structured_data["media"] = all_media
                    logger.info(f"Successfully integrated {len(all_media)} media summaries.")
            else:
                logger.warning("No media items were collected for visual summaries.")
                
        return structured_data

    def get_visual_summary(self, image_path: Path, model: str = "gemini/gemini-1.5-flash") -> str:
        """画像を個別に解析して要約文を生成する"""
        logger.info(f"Generating visual summary for {image_path.name}...")
        image_url = self._encode_image_to_base64_url(image_path)
        
        prompt = "この画像・グラフの内容を詳細に説明してください。RAGでの検索に役立つよう、構成、傾向、重要な数値、テキスト情報を含めてください。回答の冒頭に [画像概要] と付けて出力してください。"
        
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": image_url}}
                ]
            }
        ]
        
        try:
            response = litellm.completion(
                model=model,
                messages=messages,
                api_key=self.api_key,
                base_url=self.base_url
            )
            return response.choices[0].message.content or ""
        except Exception as e:
            logger.error(f"Failed to generate visual summary: {e}")
            return "[画像概要] 解析に失敗しました。"

    def extract_rag_chunks(self, excel_path: str, model: str = "gemini/gemini-1.5-flash"):
        """Excelを解析してRAG用のチャンクを生成する"""
        # 画像概要込みで構造化データを取得
        structured_data = self.extract_structured_data(Path(excel_path), model=model, include_visual_summaries=True)
        
        sheet_results = {}
        for sheet_name, sheet_data in structured_data.get("sheets", {}).items():
            
            yaml_source = ""
            if isinstance(sheet_data, dict) and "_raw_yaml" in sheet_data:
                yaml_source = sheet_data.pop("_raw_yaml")
            
            # Create isolated data payload for this sheet
            single_sheet_data = {"sheets": {sheet_name: sheet_data}}
            if "media" in structured_data: # map media starting with the sheet_name
                single_sheet_data["media"] = [m for m in structured_data["media"] if m.get("filename", "").startswith(sheet_name+"_")]
            
            markdown_text = self.rag_converter.convert(single_sheet_data)
            
            # メタデータの付与
            metadata = {
                "source_file": Path(excel_path).name,
                "sheet_name": sheet_name,
                "extraction_model": model
            }
            chunker = RagChunker(metadata=metadata)
            chunks = chunker.chunk(markdown_text, f"{Path(excel_path).name} - {sheet_name}")
            
            sheet_results[sheet_name] = {
                "chunks": chunks,
                "markdown": markdown_text,
                "structured": single_sheet_data,
                "yaml": yaml_source
            }
            
        return sheet_results, structured_data
