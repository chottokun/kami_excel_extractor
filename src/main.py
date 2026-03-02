import os
import json
import base64
from pathlib import Path
from openai import OpenAI
from extractor import extract_universal_map
from converter import convert_to_png

# 環境変数の読み込み（コンテナ内では直接OS環境変数から読み込む）
LITELLM_PROXY_URL = os.getenv("LITELLM_PROXY_URL", "http://litellm-proxy:4000")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def encode_image(image_path: Path):
    """画像をBase64にエンコードする"""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def process_excel(excel_path: Path):
    """Excelファイルを処理し、構造化JSONを生成する"""
    print(f"\nProcessing: {excel_path.name}")
    
    # 1. 画像レンダリング (LibreOffice)
    png_path = convert_to_png(excel_path, OUTPUT_DIR)
    
    # 2. 汎用座標マップ抽出 (openpyxl)
    cells_map = extract_universal_map(excel_path)
    
    # 3. VLM 呼び出し (LiteLLM Proxy)
    client = OpenAI(
        api_key="sk-1234", # litellm-config.yamlで設定したマスターキー
        base_url=f"{LITELLM_PROXY_URL}/v1"
    )
    
    base64_image = encode_image(png_path)
    
    system_prompt = """あなたはExcelドキュメントの構造化エキスパートです。
添付の「PNG画像」で視覚的なレイアウト（項目同士の空間的な関係）を把握し、提供される「座標マップ」から正確なテキスト値を引用して、構造化JSONを出力してください。
画像内のOCRは行わず、必ず座標マップ内の座標に紐づく値を使用してください。"""

    user_content = [
        {
            "type": "text",
            "text": f"""以下の座標マップを画像のガイドとして使用してください。

【座標マップ】
{json.dumps(cells_map, ensure_ascii=False)}"""
        },
        {
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{base64_image}"
            }
        },
        {
            "type": "text",
            "text": "このExcelシートを解析し、JSON形式で情報を抽出してください。"
        }
    ]

    print("Requesting VLM via LiteLLM...")
    try:
        response = client.chat.completions.create(
            model="gemini-vision",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            response_format={"type": "json_object"}
        )
        
        result_json = response.choices[0].message.content
        output_path = OUTPUT_DIR / f"{excel_path.stem}_extracted.json"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result_json)
        print(f"Extraction successful: {output_path}")
        
    except Exception as e:
        print(f"Error during VLM request: {e}")

def main():
    # 入出力ディレクトリの作成
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    # 簡易的なファイル監視（ループ）
    print(f"Monitoring {INPUT_DIR} for new files...")
    import time
    processed_files = set()
    
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed_files:
                try:
                    process_excel(f)
                    processed_files.add(f)
                except Exception as e:
                    print(f"Failed to process {f.name}: {e}")
        time.sleep(5)

if __name__ == "__main__":
    main()
