import os
import logging
from dotenv import load_dotenv
from kami_excel_extractor import KamiExcelExtractor

load_dotenv()
logging.basicConfig(level=logging.INFO)

def run_library_test():
    api_key = os.getenv("GEMINI_API_KEY")
    extractor = KamiExcelExtractor(api_key=api_key, output_dir="data/lib_test")
    
    # 既存のサンプルファイルを使用
    result = extractor.extract_structured_data("complex_report.xlsx")
    
    print("\n--- Library Extraction Result ---")
    import json
    print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    run_library_test()
