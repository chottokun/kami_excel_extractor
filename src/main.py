import os
import time
import logging
from pathlib import Path
from dotenv import load_dotenv
from kami_excel_extractor import KamiExcelExtractor

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 設定の読み込み
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def main():
    if not GEMINI_API_KEY:
        logger.error("GEMINI_API_KEY is not set.")
        return

    # ライブラリの初期化
    extractor = KamiExcelExtractor(api_key=GEMINI_API_KEY, output_dir=str(OUTPUT_DIR))
    
    logger.info(f"Library Mode Pipeline started. Monitoring {INPUT_DIR}...")
    processed = set()
    
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed:
                try:
                    logger.info(f"Processing: {f.name}")
                    # 確実に動作するモデルを指定
                    result = extractor.extract_structured_data(f, model="gemini/gemini-2.5-flash")
                    
                    # 結果の保存
                    import json
                    output_path = OUTPUT_DIR / f"{f.stem}_lib_result.json"
                    with open(output_path, "w", encoding="utf-8") as out_f:
                        json.dump(result, out_f, ensure_ascii=False, indent=2)
                    
                    logger.info(f"Success: {output_path}")
                    processed.add(f)
                except Exception as e:
                    logger.error(f"Failed to process {f.name}: {e}")
        
        time.sleep(10)

if __name__ == "__main__":
    main()
