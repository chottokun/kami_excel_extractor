import json
import logging
from pathlib import Path
import google.generativeai as genai
from .extractor import MetadataExtractor
from .converter import ExcelConverter

logger = logging.getLogger(__name__)

class KamiExcelExtractor:
    """Excelから構造化JSONを抽出するメインクラス"""
    
    def __init__(self, api_key: str, output_dir: str = "output"):
        self.api_key = api_key.strip("'\" ")
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.extractor = MetadataExtractor(self.output_dir)
        self.converter = ExcelConverter(self.output_dir)
        
        genai.configure(api_key=self.api_key)

    def _get_best_model(self):
        try:
            for m in genai.list_models():
                if 'generateContent' in m.supported_generation_methods and 'flash' in m.name.lower():
                    return m.name
        except:
            pass
        return "models/gemini-1.5-flash"

    def extract_structured_data(self, excel_path: str, system_prompt: str = None):
        excel_path = Path(excel_path)
        logger.info(f"Processing {excel_path.name}...")
        
        # 1. 物理抽出
        png_path = self.converter.convert(excel_path)
        metadata = self.extractor.extract(excel_path)
        
        # 2. VLMリクエスト
        model_name = self._get_best_model()
        model = genai.GenerativeModel(model_name)
        
        with open(png_path, "rb") as f:
            image_data = f.read()

        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。座標マップの値を正解として引用し、構造化JSONを出力してください。"

        content = [
            f"座標マップ:\n{json.dumps(metadata, ensure_ascii=False)}",
            {"mime_type": "image/png", "data": image_data},
            "回答は純粋なJSONのみを返してください。"
        ]

        response = model.generate_content([system_prompt] + content)
        result_text = response.text.strip()
        
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0].strip()
            
        return json.loads(result_text)
