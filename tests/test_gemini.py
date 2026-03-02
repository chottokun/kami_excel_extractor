import os
import json
from openai import OpenAI
from extractor import extract_universal_map

def force_load_env():
    env_path = ".env"
    if os.path.exists(env_path):
        with open(env_path, "r") as f:
            for line in f:
                if "=" in line and not line.startswith("#"):
                    k, v = line.strip().split("=", 1)
                    os.environ[k] = v.strip("'\"")
        print(f"Forced loaded {env_path}")

def test_extraction():
    force_load_env()
    
    GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
    # 明示的に gemini-1.5-flash を指定
    GEMINI_MODEL = "gemini-1.5-flash"

    if not GEMINI_API_KEY:
        print("Error: GEMINI_API_KEY not found in os.environ")
        return
    
    # 座標マップの抽出
    excel_file = "sample_hoganshi.xlsx"
    cells_map = extract_universal_map(excel_file)
    
    # プロンプト構築
    system_prompt = """あなたは、複雑なExcelドキュメントの構造化を専門とするデータアナリストです。
提供される「座標マップ（JSON形式）」を論理的に解釈し、指定されたJSONフォーマットで情報を抽出してください。
最終回答は純粋なJSONのみを返してください。"""

    user_content = f"""以下の座標マップから情報を抽出してください。

【座標マップ】
{json.dumps(cells_map, ensure_ascii=False)}

【抽出スキーマ（JSON形式）】
{{
  "report_title": "報告書のタイトル",
  "meta_info": {{
    "reporter": "報告者名",
    "date": "日付"
  }},
  "content": "報告内容の要約",
  "details": [
    {{
      "id": 1,
      "task_name": "作業名",
      "hours": 0.0
    }}
  ]
}}"""

    # v1 エンドポイントを試す
    client = OpenAI(
        api_key=GEMINI_API_KEY,
        base_url="https://generativelanguage.googleapis.com/v1beta/"
    )

    print(f"Geminiにリクエスト送信中... (Model: {GEMINI_MODEL})")
    
    try:
        # openai クライアントでの呼び出し方を調整
        # モデル名に models/ を付けないのが OpenAI 互換モードの標準
        response = client.chat.completions.create(
            model=GEMINI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ]
        )

        result = response.choices[0].message.content
        print("\n--- 抽出結果 ---")
        print(result)
    except Exception as e:
        print(f"\nError occurred: {e}")

if __name__ == "__main__":
    test_extraction()
