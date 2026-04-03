import os
import sys
import logging
from pathlib import Path
from dotenv import load_dotenv
import json

# ==============================================================================
# 共通パスと環境設定
# ==============================================================================
# プロジェクトルートの動的解決 (Docker内実行を考慮)
project_root = Path(__file__).parent.resolve()
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor

# ログの設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("run_batch_test")

def main():
    # .env ファイルの読み込み
    load_dotenv(project_root / ".env")
    
    # LLM設定の取得
    api_key = os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")
    model = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"
    
    output_dir = project_root / "data" / "output"
    input_dir = project_root / "data" / "input"
    
    # 出力ディレクトリの作成
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 抽出器の初期化
    extractor = KamiExcelExtractor(api_key=api_key, output_dir=str(output_dir))
    
    # 入力ファイルの列挙 (すべての .xlsx ファイル)
    xlsx_files = sorted(list(input_dir.glob("*.xlsx")))
    
    if not xlsx_files:
        logger.warning(f"No Excel files found in {input_dir}")
        return

    logger.info(f"Found {len(xlsx_files)} files. Starting batch processing using model: {model}")
    
    results_summary = []

    for input_file in xlsx_files:
        file_summary = {"filename": input_file.name, "status": "Pending", "error": None}
        logger.info(f"--- Processing: {input_file.name} ---")
        
        try:
            # 1. RAG チャンクの抽出
            # 実データのため、視覚的な要約 (include_visual_summaries) も含めて実行
            rag_results, augmented_data = extractor.extract_rag_chunks(input_file, model=model)
            
            # 2. 結果の保存 (JSON)
            result_path = output_dir / f"{input_file.stem}_lib_result.json"
            with open(result_path, "w", encoding="utf-8") as out_f:
                json.dump(augmented_data, out_f, ensure_ascii=False, indent=2)
            
            # 3. Markdown の保存
            md_content = "\n\n".join([res["markdown"] for res in rag_results.values()])
            md_output_path = output_dir / f"{input_file.stem}_rag.md"
            with open(md_output_path, "w", encoding="utf-8") as out_f:
                out_f.write(md_content)
            
            # 4. PDF の生成
            logger.info(f"Generating PDF for: {input_file.name}")
            pdf_path = extractor.doc_generator.generate_pdf(md_content, f"{input_file.stem}_report")
            
            if pdf_path:
                logger.info(f"Successfully generated PDF: {pdf_path.name}")
                file_summary["status"] = "Success"
            else:
                logger.warning(f"Failed to generate PDF for: {input_file.name} (Check LibreOffice setup)")
                file_summary["status"] = "Partial Success (PDF Failed)"
                
        except Exception as e:
            logger.error(f"Error processing {input_file.name}: {str(e)}")
            file_summary["status"] = "Failed"
            file_summary["error"] = str(e)
        
        results_summary.append(file_summary)
        logger.info(f"Finished: {input_file.name} (Status: {file_summary['status']})")

    # 最終的な集計レポートの出力
    print("\n" + "="*50)
    print("BATCH PROCESSING SUMMARY")
    print("="*50)
    success_count = sum(1 for s in results_summary if "Success" in s["status"])
    print(f"Total Files: {len(results_summary)}")
    print(f"Success    : {success_count}")
    print(f"Failed     : {len(results_summary) - success_count}")
    print("-"*50)
    for res in results_summary:
        err_msg = f" (Error: {res['error']})" if res['error'] else ""
        print(f"[{res['status']:<25}] {res['filename']}{err_msg}")
    print("="*50)

if __name__ == "__main__":
    main()
