import os
import json
import base64
import time
from pathlib import Path
from openai import OpenAI
from extractor import extract_comprehensive_map
from converter import convert_to_png

# 環境変数の読み込み
LITELLM_PROXY_URL = os.getenv("LITELLM_PROXY_URL", "http://litellm-proxy:4000")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def encode_image_to_base64(image_path: Path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def process_excel_litellm(excel_path: Path):
    print(f"\n--- Processing via LiteLLM: {excel_path.name} ---")
    
    # 1. 前処理（画像変換とメタデータ抽出）
    png_path = convert_to_png(excel_path, OUTPUT_DIR)
    full_map = extract_comprehensive_map(excel_path, OUTPUT_DIR)
    base64_image = encode_image_to_base64(png_path)

    # 2. LiteLLM Proxy (OpenAI互換) クライアントの初期化
    client = OpenAI(
        api_key="sk-1234", # LITELLM_MASTER_KEY
        base_url=f"{LITELLM_PROXY_URL}/v1"
    )

    system_prompt = """あなたはExcel構造化の専門家です。
画像と座標マップを照らし合わせ、指定されたJSON形式で正確に抽出してください。
画像内のOCRは行わず、座標マップの値を『正解』として引用してください。"""

    # LiteLLMの仕様に基づくOpenAI形式のマルチモーダルメッセージ
    messages = [
        {"role": "system", "content": system_prompt},
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": f"座標マップ:\n{json.dumps(full_map, ensure_ascii=False)}"
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{base64_image}"
                    }
                },
                {
                    "type": "text",
                    "text": "このExcelデータを構造化JSONで出力してください。回答はJSONのみを返してください。"
                }
            ]
        }
    ]

    print(f"Requesting via LiteLLM Proxy -> unified-vision-model...")
    max_retries = 20 # 待機時間を大幅に増加 (約200秒)
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="unified-vision-model",
                messages=messages,
                response_format={"type": "json_object"}
            )
            
            result_text = response.choices[0].message.content
            output_path = OUTPUT_DIR / f"{excel_path.stem}_structured.json"
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(result_text)
            print(f"Successfully structured: {output_path}")
            break
            
        except Exception as e:
            if ("Connection refused" in str(e) or "Connection error" in str(e)) and attempt < max_retries - 1:
                print(f"LiteLLM Proxy still starting (attempt {attempt+1}/{max_retries}). Waiting 10s...")
                time.sleep(10)
                continue
            import traceback
            print(f"LiteLLM Request Failed after {attempt+1} attempts: {e}")
            traceback.print_exc()
            break

def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Gateway Mode: Monitoring {INPUT_DIR}...")
    processed_files = set()
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed_files:
                process_excel_litellm(f)
                processed_files.add(f)
        time.sleep(10)

if __name__ == "__main__":
    main()
