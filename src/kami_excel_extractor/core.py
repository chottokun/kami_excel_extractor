import json
import logging
import base64
from pathlib import Path
import litellm
from .extractor import MetadataExtractor
from .converter import ExcelConverter

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
        
        if api_key:
            # LiteLLMは環境変数を優先するため、明示的にセット
            self.api_key = api_key.strip("'\" ")
        else:
            self.api_key = None

    def _encode_image_to_base64_url(self, image_path: Path):
        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode("utf-8")
            return f"data:image/png;base64,{encoded}"

    def extract_structured_data(self, excel_path: str, model: str = "gemini/gemini-1.5-flash", system_prompt: str = None):
        """
        Excelを解析して構造化データを取得する
        
        Args:
            excel_path: Excelファイルへのパス
            model: 使用するモデル名 (例: "openai/gpt-4o", "gemini/gemini-1.5-flash", "azure/...")
            system_prompt: カスタムの指示（任意）
        """
        excel_path = Path(excel_path)
        logger.info(f"Processing {excel_path.name} with model {model}...")
        
        # 1. 物理抽出とレンダリング
        png_path = self.converter.convert(excel_path)
        metadata = self.extractor.extract(excel_path)
        image_url = self._encode_image_to_base64_url(png_path)
        
        # 2. メッセージの構築 (OpenAI互換形式)
        if not system_prompt:
            system_prompt = "あなたはExcel構造化の専門家です。提供された座標マップの値を正解として引用し、構造化JSONを出力してください。"

        messages = [
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": f"座標マップ:\n{json.dumps(metadata, ensure_ascii=False)}"},
                    {"type": "image_url", "image_url": {"url": image_url}},
                    {"type": "text", "text": "Excelデータを構造化JSONで出力してください。回答は純粋なJSONのみを返してください。"}
                ]
            }
        ]

        # 3. LiteLLMによる呼び出し
        try:
            response = litellm.completion(
                model=model,
                messages=messages,
                api_key=self.api_key,
                base_url=self.base_url,
                response_format={"type": "json_object"}
            )
            
            result_text = response.choices[0].message.content
            if result_text is None:
                raise ValueError("VLM returned an empty response.")

            # Markdownの除去
            if "```json" in result_text:
                result_text = result_text.split("```json")[1].split("```")[0].strip()
            elif "```" in result_text:
                result_text = result_text.split("```")[1].split("```")[0].strip()
                
            return json.loads(result_text)
            
        except Exception as e:
            logger.error(f"LLM request failed: {e}")
            raise
