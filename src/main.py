import os
import json
import base64
import time
import logging
from pathlib import Path
from openai import OpenAI
from src.extractor import extract_comprehensive_map
from src.converter import convert_to_png

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 環境変数
LITELLM_PROXY_URL = os.getenv("LITELLM_PROXY_URL", "http://litellm-proxy:4000")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")
MASTER_KEY = os.getenv("LITELLM_MASTER_KEY", "sk-1234")

def encode_image_to_base64(image_path: Path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def call_vlm_via_proxy(full_map, base64_image, model_name="unified-vision-model"):
    """LiteLLM Proxy経由でVLMを呼び出す（リトライ付き）"""
    client = OpenAI(api_key=MASTER_KEY, base_url=f"{LITELLM_PROXY_URL}/v1")
    
    messages = [
        {"role": "system", "content": "あなたはExcel構造化の専門家です。提供された座標マップの値を正解として引用してください。"},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": f"座標マップ:\n{json.dumps(full_map, ensure_ascii=False)}"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}},
                {"type": "text", "text": "Excelデータを構造化JSONで出力してください。"}
            ]
        }
    ]

    max_retries = 10
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model_name,
                messages=messages,
                response_format={"type": "json_object"}
            )
            return response.choices[0].message.content
        except Exception as e:
            error_msg = str(e).lower()
            # LiteLLM起動中、または一時的なエラーの場合のみ待機
            if any(kw in error_msg for kw in ["connection", "refused", "rate_limit", "timeout"]):
                logger.warning(f"Proxy still warming up or busy (attempt {attempt+1}/{max_retries})...")
                time.sleep(15)
                continue
            logger.error(f"Fatal error during VLM request: {e}")
            raise

def process_file(excel_path: Path):
    logger.info(f"Starting processing: {excel_path.name}")
    try:
        # 1. 画像変換
        png_path = convert_to_png(excel_path, OUTPUT_DIR)
        
        # 2. メタデータ抽出
        full_map = extract_comprehensive_map(excel_path, OUTPUT_DIR)
        
        # 3. Base64化
        base64_img = encode_image_to_base64(png_path)
        
        # 4. VLMリクエスト
        result_json = call_vlm_via_proxy(full_map, base64_img)
        
        # 5. 保存
        output_path = OUTPUT_DIR / f"{excel_path.stem}_result.json"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result_json)
        logger.info(f"Successfully processed: {output_path}")
        
    except Exception as e:
        logger.error(f"Failed to process {excel_path.name}: {e}")

def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    logger.info(f"Pipeline started. Monitoring {INPUT_DIR}...")
    processed = set()
    
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed:
                process_file(f)
                processed.add(f)
        time.sleep(5)

if __name__ == "__main__":
    main()
