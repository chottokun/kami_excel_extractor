import json
import logging
import base64
from pathlib import Path
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
        
        # 2. メッセージの構築
        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。提供された座標マップの値を正解として引用し、構造化JSONを出力してください。"

        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"座標マップ:\n{json.dumps(raw_data, ensure_ascii=False)}"},
                    {"type": "image_url", "image_url": {"url": image_url}},
                    {"type": "text", "text": "Excelデータを構造化JSONで出力してください。"}
                ]
            }
        ]

        # 3. LLM呼び出し
        try:
            response = litellm.completion(
                model=model,
                messages=messages,
                api_key=self.api_key,
                base_url=self.base_url,
                response_format={"type": "json_object"}
            )
            result_text = response.choices[0].message.content
            if "```json" in result_text:
                result_text = result_text.split("```json")[1].split("```")[0].strip()
            structured_data = json.loads(result_text)
            
            # 4. 画像概要の付与 (必要に応じて)
            if include_visual_summaries:
                all_media = []
                sheets = raw_data.get("sheets", {})
                logger.info(f"raw_data sheets: {list(sheets.keys())}")
                for sheet_name, sheet_data in sheets.items():
                    media_list = sheet_data.get("media", [])
                    logger.info(f"Sheet '{sheet_name}' has {len(media_list)} media items.")
                    for media_item in media_list:
                        media_path = self.output_dir / "media" / media_item["filename"]
                        if media_path.exists():
                            logger.info(f"Found media file: {media_path}. Generating summary...")
                            summary = self.get_visual_summary(media_path, model=model)
                            media_item["visual_summary"] = summary
                            all_media.append(media_item)
                        else:
                            logger.error(f"Media file NOT FOUND: {media_path}")
                
                if all_media:
                    structured_data["media"] = all_media
                    logger.info(f"Successfully integrated {len(all_media)} media summaries.")
                else:
                    logger.warning("No media items were collected for visual summaries.")
                    
            return structured_data
        except Exception as e:
            logger.error(f"Structured extraction failed: {e}")
            raise

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
        
        markdown_text = self.rag_converter.convert(structured_data)
        
        # メタデータの付与
        metadata = {
            "source_file": Path(excel_path).name,
            "extraction_model": model
        }
        chunker = RagChunker(metadata=metadata)
        chunks = chunker.chunk(markdown_text, Path(excel_path).name)
        return chunks, markdown_text, structured_data
