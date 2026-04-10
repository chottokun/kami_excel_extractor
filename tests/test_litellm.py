import os
import logging
import json
from dotenv import load_dotenv
from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

load_dotenv()
logging.basicConfig(level=logging.INFO)

def test_litellm_gemini_integration():
    """LiteLLMライブラリを介してGeminiで抽出できるかテスト"""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        print("Skip test: GEMINI_API_KEY not found")
        return

    # ライブラリの初期化 (LiteLLMを内部で使用)
    extractor = KamiExcelExtractor(api_key=api_key, output_dir="data/test_litellm")
    
    # コンテナ内の正しいパスを使用
    excel_path = "/app/data/input/complex_report.xlsx"
    
    # 複雑なレポートを使用して抽出
    print(f"\n--- Requesting via LiteLLM (Gemini) using {excel_path} ---")
    result = extractor.extract_structured_data(
        excel_path, 
        options=ExtractionOptions(model="gemini/gemini-2.5-flash")
    )
    
    print("\n--- Extraction Result ---")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    
    # 基本的な構造チェック
    assert isinstance(result, dict)
    assert len(result.keys()) > 0

if __name__ == "__main__":
    test_litellm_gemini_integration()
