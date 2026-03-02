import os
import json
import base64
import time
from pathlib import Path
import google.generativeai as genai
from extractor import extract_comprehensive_map
from converter import convert_to_png

# 環境変数の読み込み
raw_key = os.getenv("GEMINI_API_KEY")
GEMINI_API_KEY = raw_key.strip("'\" ").strip() if raw_key else None
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def get_available_flash_model():
    """利用可能なFlashモデルを自動的に取得する"""
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods and 'flash' in m.name.lower():
                print(f"Auto-selected model: {m.name}")
                return m.name
    except Exception as e:
        print(f"Error listing models: {e}")
    return "models/gemini-1.5-flash" # フォールバック

def process_excel_complex(excel_path: Path):
    print(f"\n--- Processing Complex: {excel_path.name} ---")
    
    # 1. 総合メタデータとメディアの抽出
    try:
        full_map = extract_comprehensive_map(excel_path, OUTPUT_DIR)
        print(f"Success: Extracted metadata and media from {len(full_map['sheets'])} sheets")
    except Exception as e:
        print(f"Error during metadata extraction: {e}")
        return

    # 2. 画像レンダリング
    try:
        png_path = convert_to_png(excel_path, OUTPUT_DIR)
        print(f"Success: Full image rendered at {png_path}")
    except Exception as e:
        print(f"Error rendering image: {e}")
        return

    # 3. VLM 呼び出し (自動判別モデル)
    current_model_name = get_available_flash_model()
    model = genai.GenerativeModel(current_model_name)
    
    with open(png_path, "rb") as f:
        full_image_data = f.read()

    system_prompt = """あなたは日本の複雑なExcelドキュメントを解析する第一人者です。
複数のシート、方眼紙レイアウト、埋め込まれた現場写真を含むデータを正確に構造化してください。
回答は純粋なJSONのみを返してください。"""

    user_content = [
        f"以下の座標マップ（シート別）をガイドとして使用してください。\n\n【座標マップ】\n{json.dumps(full_map, ensure_ascii=False)}",
        {"mime_type": "image/png", "data": full_image_data},
        "全てのシートから情報を抽出し、報告書の内容と写真の説明を統合したJSONを出力してください。"
    ]

    print(f"Requesting Gemini ({current_model_name})...")
    try:
        response = model.generate_content([system_prompt] + user_content)
        result_text = response.text.strip()
        if "```json" in result_text:
            result_text = result_text.split("```json")[1].split("```")[0].strip()
        
        output_path = OUTPUT_DIR / f"{excel_path.stem}_complex_extracted.json"
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(result_text)
        print(f"Extraction successful: {output_path}")
        
    except Exception as e:
        import traceback
        print(f"Error during VLM request: {e}")
        traceback.print_exc()

def main():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Monitoring {INPUT_DIR} for new files...")
    processed_files = set()
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed_files:
                process_excel_complex(f)
                processed_files.add(f)
        time.sleep(10)

if __name__ == "__main__":
    main()
