import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import json

project_root = Path("/app")
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor

load_dotenv(project_root / ".env")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    print("GEMINI_API_KEY not set")
    sys.exit(1)

OUTPUT_DIR = project_root / "data" / "output"
INPUT_FILE = project_root / "data" / "input" / "gattai_matrix_20260214.xlsx"

extractor = KamiExcelExtractor(api_key=GEMINI_API_KEY, output_dir=str(OUTPUT_DIR))
print(f"Processing: {INPUT_FILE.name}")

rag_chunks, md_content, augmented_data = extractor.extract_rag_chunks(INPUT_FILE, model="gemini/gemini-2.5-flash")

result_path = OUTPUT_DIR / f"{INPUT_FILE.stem}_lib_result.json"
with open(result_path, "w", encoding="utf-8") as out_f:
    json.dump(augmented_data, out_f, ensure_ascii=False, indent=2)

rag_output_path = OUTPUT_DIR / f"{INPUT_FILE.stem}_rag_chunks.json"
with open(rag_output_path, "w", encoding="utf-8") as out_f:
    json.dump(rag_chunks, out_f, ensure_ascii=False, indent=2)

md_output_path = OUTPUT_DIR / f"{INPUT_FILE.stem}_rag.md"
with open(md_output_path, "w", encoding="utf-8") as out_f:
    out_f.write(md_content)

print("generating PDF report ...")
extractor.doc_generator.generate_pdf(md_content, f"{INPUT_FILE.stem}_report")

print("Done. Generated files:")
print(f" - {result_path}")
print(f" - {rag_output_path}")
print(f" - {md_output_path}")
print(f" - {OUTPUT_DIR / f'{INPUT_FILE.stem}_report.pdf'}")
