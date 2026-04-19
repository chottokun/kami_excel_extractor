import os
import time
import logging
import json
from pathlib import Path
from dotenv import load_dotenv
from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import RagOptions
from kami_excel_extractor.utils import secure_filename

# ロギング設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 設定の読み込み
load_dotenv()
INPUT_DIR = Path("data/input")
OUTPUT_DIR = Path("data/output")

# テストのためにグローバル変数として定義
LLM_API_KEY = os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")
LLM_MODEL = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini/gemini-1.5-flash"
LLM_TIMEOUT = float(os.getenv("LLM_TIMEOUT") or 1800.0)

def _save_sheet_results(sheet_name: str, res: dict, target_dir: Path, extractor: KamiExcelExtractor) -> None:
    """1つのシートの抽出結果を各種フォーマットで保存する。"""
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

def process_file(f: Path, extractor: KamiExcelExtractor, model: str) -> bool:
    """
    1つのExcelファイルを処理し、結果を保存する。
    成功した場合はTrue、失敗した場合はFalseを返す。
    """
    try:
        logger.info(f"Processing: {f.name}")
        # 解析の実行（画像概要生成を含む）
        rag_options = RagOptions(model=model)
        sheet_results, full_structured_data = extractor.extract_rag_chunks(f, options=rag_options)

        # 安全なディレクトリ名の作成
        safe_stem = secure_filename(f.stem)
        target_dir = OUTPUT_DIR / safe_stem
        target_dir.mkdir(parents=True, exist_ok=True)

        # 構造化された抽出結果全体（参考用）
        full_result_path = target_dir / "full_lib_result.json"
        with open(full_result_path, "w", encoding="utf-8") as out_f:
            json.dump(full_structured_data, out_f, ensure_ascii=False, indent=2)

        for sheet_name, res in sheet_results.items():
            _save_sheet_results(sheet_name, res, target_dir, extractor)

        logger.info(f"Success: Outputs saved to {target_dir}")
        return True
    except Exception as e:
        logger.error(f"Failed to process {f.name}: {e}")
        return False

def _process_pending_files(extractor: KamiExcelExtractor, model: str, processed: set[Path]) -> None:
    """入力ディレクトリを監視し、未処理のファイルを処理する。"""
    files = list(INPUT_DIR.glob("*.xlsx"))
    for f in files:
        if f not in processed:
            if process_file(f, extractor, model):
                processed.add(f)

def main():
    # 実行時に最新の値を取得 (テストパッチ対応)
    global LLM_API_KEY, LLM_MODEL, LLM_TIMEOUT
    api_key = LLM_API_KEY or os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")
    model = LLM_MODEL or os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini/gemini-1.5-flash"
    timeout = LLM_TIMEOUT or float(os.getenv("LLM_TIMEOUT") or 1800.0)

    if not api_key and "ollama" not in model:
        logger.error("LLM_API_KEY or GEMINI_API_KEY is not set.")
        return

    # ライブラリの初期化
    extractor = KamiExcelExtractor(api_key=api_key, output_dir=str(OUTPUT_DIR), timeout=timeout)
    
    logger.info(f"Library Mode Pipeline started. Monitoring {INPUT_DIR}...")
    processed = set()
    
    while True:
        _process_pending_files(extractor, model, processed)
        time.sleep(10)

if __name__ == "__main__":
    main()
