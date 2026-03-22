import os
import sys
from pathlib import Path
import json
import time

project_root = Path("/app")
sys.path.append(str(project_root / "src"))

from kami_excel_extractor.extractor import MetadataExtractor

INPUT_FILE = project_root / "data" / "input" / "gattai_matrix_20260214.xlsx"
OUTPUT_DIR = project_root / "data" / "output"

extractor = MetadataExtractor(output_dir=OUTPUT_DIR)
print("Extracting raw html tables...")
t0 = time.time()
raw_data = extractor.extract(INPUT_FILE)
t1 = time.time()
print(f"Extraction took {t1 - t0:.2f} seconds.")

# Check lengths
sheets = list(raw_data["sheets"].keys())
print(f"Sheets: {sheets}")
total_len = 0
for s in sheets:
    html_str = raw_data["sheets"][s].get("html", "")
    length = len(html_str)
    total_len += length
    print(f"Sheet {s}: {length} HTML chars, ~{length//4} tokens.")

print(f"Total HTML string length: {total_len} characters.")
print(f"Total Estimated token count (char / 4): {total_len // 4}")
