import os
import json
import google.generativeai as genai
from extractor import extract_universal_map

def force_load_env():
    env_path = ".env"
    if os.path.exists(env_path):
        with open(env_path, "r") as f:
            for line in f:
                if "=" in line and not line.startswith("#"):
                    k, v = line.strip().split("=", 1)
                    os.environ[k] = v.strip("'\"")

def test_google_sdk():
    force_load_env()
    
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    # .env のモデル名 (models/gemini-3-flash-preview) を使用
    GEMINI_MODEL = os.getenv("GEMINI_MODEL", "models/gemini-2.0-flash")
    if not GEMINI_MODEL.startswith("models/"):
        GEMINI_MODEL = f"models/{GEMINI_MODEL}"

    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(GEMINI_MODEL)

    # 1. 座標マップの抽出
    excel_file = "sample_hoganshi.xlsx"
    cells_map = extract_universal_map(excel_file)
    
    # 2. プロンプト構築
    prompt = f"""あなたはExcelドキュメントの構造化エキスパートです。
以下のJSON形式の「座標マップ」は、Excelの各セルの値、座標、結合状態、罫線情報を含んでいます。
方眼紙形式や報告書形式のレイアウトを論理的に解釈し、構造化されたJSONとして抽出してください。

【座標マップ】
{json.dumps(cells_map, ensure_ascii=False)}

【抽出スキーマ】
{{
  "report_title": "報告書のタイトル",
  "meta_info": {{
    "reporter": "報告者名",
    "date": "日付"
  }},
  "content": "報告内容の全文",
  "details": [
    {{
      "id": 1,
      "task_name": "作業名",
      "hours": 0.0
    }}
  ]
}}

回答は純粋なJSONのみを返してください。"""

    print(f"Gemini (Google SDK) にリクエスト送信中... (Model: {GEMINI_MODEL})")
    
    try:
        response = model.generate_content(prompt)
        print("\n--- 抽出結果 ---")
        print(response.text)
    except Exception as e:
        print(f"\nError occurred: {e}")

if __name__ == "__main__":
    test_google_sdk()
