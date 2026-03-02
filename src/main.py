import os
import json
import base64
import time
import logging
from pathlib import Path
import google.generativeai as genai
from src.extractor import extract_comprehensive_map
from src.converter import convert_to_png

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 環境変数
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip("'\" ")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def get_best_model():
    """利用可能なFlashモデルを自動判別する"""
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods and 'flash' in m.name.lower():
                return m.name
    except Exception as e:
        logger.warning(f"Model listing failed: {e}")
    return "models/gemini-1.5-flash" # フォールバック

def call_vlm_direct(full_map, image_path):
    """Gemini SDKを直接叩いてVLMを呼び出す"""
    model_name = get_best_model()
    logger.info(f"Using model: {model_name}")
    
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(model_name)
    
    with open(image_path, "rb") as f:
        image_data = f.read()

    system_prompt = """あなたはExcel構造化の専門家です。提供された座標マップの値を正解として引用してください。
埋め込まれた写真の内容（不具合箇所など）も周辺テキストと関連付けて説明に含めてください。"""

    user_content = [
        f"座標マップ:\n{json.dumps(full_map, ensure_ascii=False)}",
        {"mime_type": "image/png", "data": image_data},
        "Excelデータを構造化JSONで出力してください。回答は純粋なJSONのみを返してください。"
    ]

    try:
        response = model.generate_content([system_prompt] + user_content)
        result_text = response.text.strip()
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0].strip()
        return result_text
    except Exception as e:
        logger.error(f"VLM Direct Request Failed: {e}")
        return None

def process_file(excel_path: Path):
    logger.info(f"Starting final processing: {excel_path.name}")
    try:
        png_path = convert_to_png(excel_path, OUTPUT_DIR)
        full_map = extract_comprehensive_map(excel_path, OUTPUT_DIR)
        result_json = call_vlm_direct(full_map, png_path)
        
        if result_json:
            output_path = OUTPUT_DIR / f"{excel_path.stem}_final_result.json"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result_json)
            logger.info(f"Successfully processed: {output_path}")
        else:
            logger.error(f"Failed to get valid JSON for {excel_path.name}")
    except Exception as e:
        logger.error(f"Failed to process {excel_path.name}: {e}")

def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("Ultimate Pipeline started...")
    processed = set()
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed:
                process_file(f)
                processed.add(f)
        time.sleep(10)

if __name__ == "__main__":
    main()
