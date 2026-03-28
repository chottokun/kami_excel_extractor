import os
import time
import logging
from pathlib import Path
from dotenv import load_dotenv
from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.utils import secure_filename

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 設定の読み込み
load_dotenv()
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

def main():
    llm_api_key = os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")
    llm_model = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"

    if not llm_api_key and "ollama" not in llm_model:
        logger.error("Neither LLM_API_KEY nor GEMINI_API_KEY is set (and not using Ollama).")
        return

    # ライブラリの初期化
    extractor = KamiExcelExtractor(api_key=llm_api_key, output_dir=str(OUTPUT_DIR))
    
    logger.info(f"Library Mode Pipeline started. Monitoring {INPUT_DIR}...")
    processed = set()
    
    while True:
        files = list(INPUT_DIR.glob("*.xlsx"))
        for f in files:
            if f not in processed:
                try:
                    logger.info(f"Processing: {f.name}")
                    # 解析の実行（画像概要生成を含む）
                    sheet_results, full_structured_data = extractor.extract_rag_chunks(f, model=llm_model)
                    import json
                    
                    target_dir = OUTPUT_DIR / f.stem
                    target_dir.mkdir(parents=True, exist_ok=True)
                    
                    # 構造化された抽出結果全体（参考用）
                    full_result_path = target_dir / "full_lib_result.json"
                    with open(full_result_path, "w", encoding="utf-8") as out_f:
                        json.dump(full_structured_data, out_f, ensure_ascii=False, indent=2)
                    
                    for sheet_name, res in sheet_results.items():
                        # 安全なファイル名の作成
                        safe_sheet_name = secure_filename(sheet_name)
                        
                        sheet_struct_path = target_dir / f"{safe_sheet_name}_lib_result.json"
                        with open(sheet_struct_path, "w", encoding="utf-8") as out_f:
                            json.dump(res["structured"], out_f, ensure_ascii=False, indent=2)
                        
                        sheet_yaml_path = target_dir / f"{safe_sheet_name}_lib_result.yaml"
                        with open(sheet_yaml_path, "w", encoding="utf-8") as out_f:
                            out_f.write(res["yaml"])
                        
                        sheet_rag_path = target_dir / f"{safe_sheet_name}_rag_chunks.json"
                        with open(sheet_rag_path, "w", encoding="utf-8") as out_f:
                            json.dump(res["chunks"], out_f, ensure_ascii=False, indent=2)
                        
                        sheet_md_path = target_dir / f"{safe_sheet_name}_rag.md"
                        with open(sheet_md_path, "w", encoding="utf-8") as out_f:
                            out_f.write(res["markdown"])
                        
                        logger.info(f"Generating PDF report for sheet {sheet_name}...")
                        extractor.doc_generator.generate_pdf(res["markdown"], f"{target_dir.name}/{safe_sheet_name}_report")
                    
                    logger.info(f"Success: Outputs saved to {target_dir}")
                    processed.add(f)
                except Exception as e:
                    logger.error(f"Failed to process {f.name}: {e}")
        
        time.sleep(10)

if __name__ == "__main__":
    main()
