import openpyxl
import json
import os
from pathlib import Path
from openpyxl.utils import get_column_letter
import io

class MetadataExtractor:
    """Excelからメタデータとメディアを抽出するクラス"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.media_dir = self.output_dir / "media"
        self.media_dir.mkdir(parents=True, exist_ok=True)

    def _get_border_info(self, cell):
        borders = {}
        if cell.border:
            if cell.border.left.style: borders["L"] = cell.border.left.style
            if cell.border.right.style: borders["R"] = cell.border.right.style
            if cell.border.top.style: borders["T"] = cell.border.top.style
            if cell.border.bottom.style: borders["B"] = cell.border.bottom.style
        return borders

    def _extract_media(self, ws, sheet_name):
        media_info = []
        if not hasattr(ws, "_images"):
            return media_info

        for idx, img in enumerate(ws._images):
            row = img.anchor._from.row + 1
            col = img.anchor._from.col + 1
            coord = f"{get_column_letter(col)}{row}"
            image_filename = f"{sheet_name}_img_{coord}_{idx}.png"
            save_path = self.media_dir / image_filename
            
            try:
                data = img.ref.read() if hasattr(img.ref, "read") else img.ref.getvalue()
                with open(save_path, "wb") as f:
                    f.write(data)
                media_info.append({"coord": coord, "filename": str(image_filename), "type": "image"})
            except Exception as e:
                pass # ロギング等
        return media_info

    def extract(self, excel_path: Path):
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        full_map = { "sheets": {} }
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            merged_ranges = [str(r) for r in ws.merged_cells.ranges]
            cells_map = []
            
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r, column=c)
                    if cell.value is None and not cell.border and not cell.fill: continue
                    
                    cell_info = {"c": cell.coordinate, "v": str(cell.value) if cell.value is not None else ""}
                    if cell.fill and cell.fill.start_color.index != '00000000':
                        cell_info["bg"] = str(cell.fill.start_color.index)
                    borders = self._get_border_info(cell)
                    if borders: cell_info["b"] = borders
                    
                    for m_range in merged_ranges:
                        if cell.coordinate in openpyxl.worksheet.cell_range.CellRange(m_range):
                            cell_info["m"] = m_range
                            break
                    cells_map.append(cell_info)
            
            full_map["sheets"][sheet_name] = {
                "cells": cells_map,
                "media": self._extract_media(ws, sheet_name)
            }
        return full_map
