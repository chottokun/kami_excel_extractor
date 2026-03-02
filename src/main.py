import os
import json
import base64
import time
from pathlib import Path
import google.generativeai as genai
from extractor import extract_universal_map
from converter import convert_to_png

# 環境変数の読み込み
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "models/gemini-2.0-flash")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def encode_image(image_path: Path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def process_excel(excel_path: Path):
    print(f"\n--- Processing: {excel_path.name} ---")
    
    # 1. 画像レンダリング (PDF -> PNG 2段階変換)
    try:
        png_path = convert_to_png(excel_path, OUTPUT_DIR)
        print(f"Success: Image rendered at {png_path}")
    except Exception as e:
        print(f"Error rendering image: {e}")
        return

    # 2. 汎用座標マップ抽出 (openpyxl)
    cells_map = extract_universal_map(excel_path)
    print(f"Success: XML map extracted ({len(cells_map)} cells)")
    
    # 3. VLM 呼び出し (Direct Google SDK)
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(GEMINI_MODEL)
    
    with open(png_path, "rb") as f:
        image_data = f.read()

    system_prompt = """あなたはExcelドキュメントの構造化エキスパートです。
添付の「PNG画像」で視覚的なレイアウト（項目同士の空間的な関係）を把握し、提供される「座標マップ」から正確なテキスト値を引用して、構造化JSONを出力してください。
画像内のOCRは行わず、必ず座標マップ内の座標に紐づく値を使用して、方眼紙・報告書形式の意図を汲み取ってください。"""

    user_content = [
        f"以下の座標マップをガイドとして使用してください。\n\n【座標マップ】\n{json.dumps(cells_map, ensure_ascii=False)}",
        {"mime_type": "image/png", "data": image_data},
        "このExcelシートを解析し、JSON形式で情報を抽出してください。回答は純粋なJSONのみを返してください。"
    ]

    print(f"Requesting Gemini ({GEMINI_MODEL})...")
    try:
        response = model.generate_content([system_prompt] + user_content)
        
        # 抽出結果の保存
        result_text = response.text.strip()
        # JSONブロックがマークダウンで囲まれている場合を考慮
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0].strip()
        
        output_path = OUTPUT_DIR / f"{excel_path.stem}_extracted.json"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result_text)
        print(f"Extraction successful: {output_path}")
        
    except Exception as e:
        print(f"Error during VLM request: {e}")

def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    
    print(f"Monitoring {INPUT_DIR} for new files...")
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
        time.sleep(10)

if __name__ == "__main__":
    main()
