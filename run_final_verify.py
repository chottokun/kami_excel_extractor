import os
import sys
import logging
import asyncio
from pathlib import Path
from dotenv import load_dotenv
import json

# プロジェクトルートの動的解決
project_root = Path(__file__).parent.resolve()
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

# ログの設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger("run_final_verify")

async def verify_file(extractor, input_file, model, output_dir):
    """1つのファイルに対して網羅的な検証を行う"""
    summary = {"filename": input_file.name, "steps": {}}
    logger.info(f"=== Testing File: {input_file.name} ===")
    
    try:
        # Step 1: 構造化抽出 (Default Mode)
        logger.info(f"Step 1: Extracting structured data (Default)")
        options = ExtractionOptions(model=model, include_visual_summaries=True)
        result = await extractor.aextract_structured_data(input_file, options=options)
        summary["steps"]["structured_extraction"] = "Success"
        
        with open(output_dir / f"{input_file.stem}_structured.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        # Step 2: RAG チャンク生成
        logger.info(f"Step 2: Generating RAG chunks")
        rag_options = RagOptions(model=model)
        rag_results, _ = await extractor.aextract_rag_chunks(input_file, options=rag_options)
        summary["steps"]["rag_chunks"] = "Success"
        
        md_content = "\n\n".join([res["markdown"] for res in rag_results.values()])
        with open(output_dir / f"{input_file.stem}_rag.md", "w", encoding="utf-8") as f:
            f.write(md_content)

        # Step 3: PDF レポート生成
        logger.info(f"Step 3: Generating PDF report")
        # セキュアな生成をテスト
        pdf_path = extractor.doc_generator.generate_pdf(md_content, f"{input_file.stem}_verify_report")
        if pdf_path and pdf_path.exists():
            summary["steps"]["pdf_generation"] = "Success"
            logger.info(f"PDF generated: {pdf_path.name}")
        else:
            summary["steps"]["pdf_generation"] = "Failed"
            logger.error("PDF generation failed")

    except Exception as e:
        logger.error(f"Error during verification of {input_file.name}: {e}", exc_info=True)
        summary["error"] = str(e)
        
    return summary

async def main():
    load_dotenv(project_root / ".env")
    
    # LLM設定
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL") or "gemini-1.5-flash"
    
    output_dir = project_root / "test_verify_out"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    extractor = KamiExcelExtractor(api_key=api_key, output_dir=str(output_dir))
    
    # 検証対象の重要ファイル
    target_files = [
        "complex_report.xlsx",
        "sample_hoganshi.xlsx",
        "data/input/final_test_media.xlsx",
        "data/input/ui_test_sample.xlsx"
    ]
    
    xlsx_files = []
    for f in target_files:
        p = Path(f)
        if not p.is_absolute():
            p = project_root / p
        if p.exists():
            xlsx_files.append(p)
        else:
            logger.warning(f"File not found: {f}")

    if not xlsx_files:
        logger.error("No test files found.")
        return

    logger.info(f"Starting final verification for {len(xlsx_files)} files using model: {model}")
    
    all_summaries = []
    for input_file in xlsx_files:
        summary = await verify_file(extractor, input_file, model, output_dir)
        all_summaries.append(summary)

    # 最終レポート
    print("\n" + "="*60)
    print("FINAL VERIFICATION REPORT")
    print("="*60)
    for s in all_summaries:
        status_line = f"File: {s['filename']}"
        print(status_line)
        if "error" in s:
            print(f"  Overall Status: FAILED")
            print(f"  Error: {s['error']}")
        else:
            for step, status in s["steps"].items():
                print(f"  - {step:<25}: {status}")
        print("-" * 60)

if __name__ == "__main__":
    asyncio.run(main())
