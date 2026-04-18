import os
import sys
import logging
import asyncio
from pathlib import Path
from dotenv import load_dotenv
import json
import time

# プロジェクトルートの動的解決
project_root = Path(__file__).parent.resolve()
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

# ログの設定
logging.basicConfig(
    level=logging.INFO, 
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler("container_full_verify.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("full_verify")

async def process_single_file(extractor, input_file, model, output_dir):
    """1つのファイルに対して全工程を直列に実行"""
    start_time = time.time()
    logger.info(f">>> START: {input_file.name}")
    
    summary = {
        "filename": input_file.name,
        "size_kb": round(input_file.stat().st_size / 1024, 2),
        "steps": {},
        "duration": 0
    }
    
    try:
        # Step 1: 構造化抽出 (画像解析・リトライ込み)
        # タイムアウトを1200秒(20分)に設定
        options = ExtractionOptions(
            model=model, 
            include_visual_summaries=True,
            use_visual_context=True
        )
        # extractor.timeout を直接変更（もしプロパティがあれば）
        extractor.timeout = 1200.0
        
        result = await extractor.aextract_structured_data(input_file, options=options)
        summary["steps"]["extraction"] = "SUCCESS"
        
        # Step 2: RAGチャンク生成
        rag_options = RagOptions(model=model)
        rag_results, _ = await extractor.aextract_rag_chunks(input_file, options=rag_options)
        summary["steps"]["rag"] = "SUCCESS"
        
        # 保存
        out_base = output_dir / input_file.stem
        out_base.mkdir(parents=True, exist_ok=True)
        
        with open(out_base / "structured.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        md_content = "\n\n".join([res["markdown"] for res in rag_results.values()])
        with open(out_base / "content.md", "w", encoding="utf-8") as f:
            f.write(md_content)

        # Step 3: PDF レポート生成 (LibreOfficeの検証)
        logger.info(f"Generating PDF for {input_file.name}...")
        pdf_path = extractor.doc_generator.generate_pdf(md_content, f"{input_file.stem}_report")
        if pdf_path and pdf_path.exists():
            summary["steps"]["pdf"] = "SUCCESS"
            logger.info(f"PDF SUCCESS: {pdf_path.name}")
        else:
            summary["steps"]["pdf"] = "FAILED"
            logger.warning(f"PDF FAILED for {input_file.name}")

    except Exception as e:
        logger.error(f"FAILED: {input_file.name} - {str(e)}", exc_info=True)
        summary["error"] = str(e)
    
    summary["duration"] = round(time.time() - start_time, 2)
    logger.info(f"<<< FINISH: {input_file.name} in {summary['duration']}s")
    return summary

async def main():
    load_dotenv(project_root / ".env")
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL") or "gemini-1.5-flash"
    
    # 全Excelファイルの収集
    input_dir = project_root / "data" / "input"
    xlsx_files = sorted([f for f in input_dir.rglob("*.xlsx") if "~$" not in f.name])
    
    output_dir = project_root / "data" / "output" / "verify_results"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 抽出器の初期化
    extractor = KamiExcelExtractor(api_key=api_key, output_dir=str(output_dir))
    
    logger.info(f"Starting Exhaustive Container Test on {len(xlsx_files)} files using {model}")
    logger.info("Running files SERIALLY to ensure stability and LLM reliability.")
    
    all_summaries = []
    for f in xlsx_files:
        # 1ファイルずつ処理
        summary = await process_single_file(extractor, f, model, output_dir)
        all_summaries.append(summary)
    
    # 最終レポート
    print("\n" + "="*90)
    print(f"{'FILENAME':<35} | {'EXTRACTION':<10} | {'PDF':<10} | {'TIME(s)':>8} | {'STATUS'}")
    print("-" * 90)
    for s in all_summaries:
        ext_status = s["steps"].get("extraction", "FAIL")
        pdf_status = s["steps"].get("pdf", "FAIL")
        overall = "OK" if "error" not in s else "FAIL"
        print(f"{s['filename'][:35]:<35} | {ext_status:<10} | {pdf_status:<10} | {s['duration']:>8} | {overall}")
    
    success_count = sum(1 for s in all_summaries if "error" not in s)
    print("="*90)
    print(f"TOTAL: {len(all_summaries)} | SUCCESS: {success_count} | FAILED: {len(all_summaries) - success_count}")

if __name__ == "__main__":
    asyncio.run(main())
