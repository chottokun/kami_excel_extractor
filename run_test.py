import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import json

project_root = Path(__file__).parent
sys.path.append(str(project_root / "src"))

from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import RagOptions

load_dotenv(project_root / ".env")
API_KEY = os.getenv("LLM_API_KEY") or os.getenv("GEMINI_API_KEY")
MODEL = os.getenv("LLM_MODEL") or os.getenv("GEMINI_MODEL") or "gemini-1.5-flash"
if not API_KEY and not os.getenv("LLM_BASE_URL") and not os.getenv("OLLAMA_BASE_URL"):
    print("Warning: API_KEY or custom BASE_URL not set. Extraction might fail depending on the model.")

OUTPUT_DIR = project_root / "data" / "output"
INPUT_FILE = project_root / "data" / "input" / "complex_report.xlsx"

extractor = KamiExcelExtractor(api_key=API_KEY, output_dir=str(OUTPUT_DIR))
print(f"Processing: {INPUT_FILE.name} with model {MODEL}")

rag_results, augmented_data = extractor.extract_rag_chunks(INPUT_FILE, options=RagOptions(model=MODEL))

# 全シートのMarkdownを統合
md_content = "\n\n".join([res["markdown"] for res in rag_results.values()])
# 全シートのチャンクを統合
rag_chunks = []
for res in rag_results.values():
    rag_chunks.extend(res["chunks"])

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
