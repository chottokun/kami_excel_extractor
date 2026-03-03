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
                    # 解析の実行（画像概要生成を含む）
                    rag_chunks, md_content, augmented_data = extractor.extract_rag_chunks(f, model="gemini/gemini-2.5-flash")
                    
                    # 構造化された抽出結果(JSON)の保存
                    import json
                    result_path = OUTPUT_DIR / f"{f.stem}_lib_result.json"
                    with open(result_path, "w", encoding="utf-8") as out_f:
                        json.dump(augmented_data, out_f, ensure_ascii=False, indent=2)
                    
                    # RAG用チャンクの保存
                    rag_output_path = OUTPUT_DIR / f"{f.stem}_rag_chunks.json"
                    with open(rag_output_path, "w", encoding="utf-8") as out_f:
                        json.dump(rag_chunks, out_f, ensure_ascii=False, indent=2)
                    
                    # Markdown形式の保存（画像概要付き）
                    md_output_path = OUTPUT_DIR / f"{f.stem}_rag.md"
                    with open(md_output_path, "w", encoding="utf-8") as out_f:
                        out_f.write(md_content)
                    
                    # PDFレポートの生成
                    logger.info(f"Generating PDF report for {f.name}...")
                    extractor.doc_generator.generate_pdf(md_content, f"{f.stem}_report")

                    logger.info(f"Success: {result_path}, {rag_output_path}, and PDF report")
                    processed.add(f)
                except Exception as e:
                    logger.error(f"Failed to process {f.name}: {e}")
        
        time.sleep(10)

if __name__ == "__main__":
    main()
