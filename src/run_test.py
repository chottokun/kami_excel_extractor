import os
import sys
import logging
from pathlib import Path
import json

logging.basicConfig(level=logging.INFO, stream=sys.stdout, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

project_root = Path("/app")
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor

# dotenv won't be in /app/.env if not mounted, but GEMINI_API_KEY is provided via docker-compose env_file
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("GEMINI_API_KEY not set")
    sys.exit(1)

OUTPUT_DIR = project_root / "data" / "output"
INPUT_FILE = project_root / "data" / "input" / "gattai_matrix_20260214.xlsx"

extractor = KamiExcelExtractor(api_key=GEMINI_API_KEY, output_dir=str(OUTPUT_DIR))
print(f"Processing: {INPUT_FILE.name}")

MODEL_ENV = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
if not MODEL_ENV.startswith("gemini/"):
    MODEL_ENV = f"gemini/{MODEL_ENV}"
print(f"Using model: {MODEL_ENV}")

sheet_results, full_structured_data = extractor.extract_rag_chunks(INPUT_FILE, model=MODEL_ENV)

target_dir = OUTPUT_DIR / INPUT_FILE.stem
target_dir.mkdir(parents=True, exist_ok=True)

for sheet_name, res in sheet_results.items():
    safe_sheet_name = sheet_name.replace("/", "_").replace("\\", "_")
    
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
    
    print(f"Generating PDF report for sheet {sheet_name}...")
    extractor.doc_generator.generate_pdf(res["markdown"], f"{target_dir.name}/{safe_sheet_name}_report")

print(f"Done. Outputs generated in {target_dir}")
