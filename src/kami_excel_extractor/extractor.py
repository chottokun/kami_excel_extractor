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

    def is_simple_table(self, ws) -> bool:
        """シートが単純な表（ヘッダー＋データ）であるかを簡易判定する"""
        # 結合セルがある場合は単純な表ではないとみなす（Kamiの出番）
        if ws.merged_cells.ranges:
            return False

        # 値が入っているセルの密度や配置をチェック
        # ここでは簡易的に、1行目に複数の値があり、それ以降も連続してデータがある場合にTrueとする
        max_r = ws.max_row
        max_c = ws.max_column

        if max_r < 2 or max_c < 1:
            return False

        # ヘッダー候補（1行目）のチェック
        header_values = [ws.cell(row=1, column=c).value for c in range(1, max_c + 1) if ws.cell(row=1, column=c).value is not None]
        if len(header_values) < 2:
            return False

        return True

    def extract_simple_table(self, ws) -> List[Dict[str, Any]]:
        """単純な表をリスト形式の辞書として抽出する"""
        data = []
        max_r = ws.max_row
        max_c = ws.max_column

        # 1行目をヘッダーとする
        headers = [str(ws.cell(row=1, column=c).value or f"Column{c}") for c in range(1, max_c + 1)]

        for r in range(2, max_r + 1):
            row_dict = {}
            has_value = False
            for c in range(1, max_c + 1):
                val = ws.cell(row=r, column=c).value
                if val is not None:
                    has_value = True
                row_dict[headers[c-1]] = str(val) if val is not None else ""
            if has_value:
                data.append(row_dict)
        return data

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
                "media": self._extract_media(ws, sheet_name),
                "is_simple": self.is_simple_table(ws)
            }
            if full_map["sheets"][sheet_name]["is_simple"]:
                full_map["sheets"][sheet_name]["structured_data"] = self.extract_simple_table(ws)
        return full_map
