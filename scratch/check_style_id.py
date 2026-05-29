import openpyxl
from pathlib import Path

wb = openpyxl.Workbook()
ws = wb.active
cell = ws["A1"]
print(f"Style ID: {getattr(cell, 'style_id', 'Not found')}")
