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
        logging.FileHandler("full_batch_test.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("full_verify")

async def process_single_file(extractor, input_file, model, output_dir):
    """1つのファイルに対して全工程を実行"""
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
        options = ExtractionOptions(model=model, include_visual_summaries=True)
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

    except Exception as e:
        logger.error(f"FAILED: {input_file.name} - {str(e)}")
        summary["error"] = str(e)
    
    summary["duration"] = round(time.time() - start_time, 2)
    logger.info(f"<<< FINISH: {input_file.name} in {summary['duration']}s")
    return summary

async def main():
    load_dotenv(project_root / ".env")
    api_key = os.getenv("LLM_API_KEY")
    model = os.getenv("LLM_MODEL") or "gemini-1.5-flash"
    
    # 全Excelファイルの収集
    xlsx_files = list(project_root.rglob("*.xlsx"))
    # 重複除外と整理
    xlsx_files = sorted(list(set([f for f in xlsx_files if "~$" not in f.name])))
    
    output_dir = project_root / "test_full_results"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 並列度を考慮した初期化 (Semaphoreは extractor 内部で制御される)
    extractor = KamiExcelExtractor(api_key=api_key, output_dir=str(output_dir))
    
    logger.info(f"Starting exhaustive test on {len(xlsx_files)} files using {model}")
    
    # 完全に非同期で一斉にスケジュール（内部のSemaphoreで流量制御される）
    tasks = [process_single_file(extractor, f, model, output_dir) for f in xlsx_files]
    summaries = await asyncio.gather(*tasks)
    
    # レポート生成
    print("\n" + "="*80)
    print(f"{'FILENAME':<40} | {'SIZE(KB)':>10} | {'TIME(s)':>8} | {'STATUS'}")
    print("-" * 80)
    for s in summaries:
        status = "OK" if "error" not in s else "FAIL"
        print(f"{s['filename'][:40]:<40} | {s['size_kb']:>10} | {s['duration']:>8} | {status}")
    
    success_count = sum(1 for s in summaries if "error" not in s)
    print("="*80)
    print(f"TOTAL: {len(summaries)} | SUCCESS: {success_count} | FAILED: {len(summaries) - success_count}")
    print(f"Results saved to: {output_dir}")

if __name__ == "__main__":
    asyncio.run(main())
