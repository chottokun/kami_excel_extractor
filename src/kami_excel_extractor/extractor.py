import openpyxl
import json
import os
from pathlib import Path
from openpyxl.utils import get_column_letter
import io
from typing import List, Dict, Any
from datetime import date, datetime
from PIL import Image

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
                # 生のバイナリを取得
                raw_data = img.ref.read() if hasattr(img.ref, "read") else img.ref.getvalue()
                
                # Pillowを使って画像として読み込み、PNGとして再保存
                # これによりEMF/WMF等の互換性問題を解決し、標準的なPNGヘッダーを付与する
                with Image.open(io.BytesIO(raw_data)) as pillow_img:
                    # RGBに変換（透過情報の扱いや、特殊な色空間の回避）
                    if pillow_img.mode in ("RGBA", "P"):
                        pillow_img = pillow_img.convert("RGB")
                    pillow_img.save(save_path, "PNG")
                
                media_info.append({"coord": coord, "filename": str(image_filename), "type": "image"})
            except Exception as e:
                # 変換不能な形式（メタファイル以外等）はスキップ
                pass 
        return media_info

    def is_simple_table(self, ws) -> bool:
        if ws.merged_cells.ranges: return False
        max_r, max_c = ws.max_row, ws.max_column
        if max_r < 2 or max_c < 1: return False
        first_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=max_c, values_only=True), [])
        header_values = [val for val in first_row if val is not None]
        return len(header_values) >= 2

    def extract_simple_table(self, ws) -> List[Dict[str, Any]]:
        data = []
        max_r, max_c = ws.max_row, ws.max_column
        header_row = next(ws.iter_rows(min_row=1, max_row=1, min_col=1, max_col=max_c, values_only=True), [])
        headers = [str(val or f"Column{i+1}") for i, val in enumerate(header_row)]

        for row in ws.iter_rows(min_row=2, max_row=max_r, min_col=1, max_col=max_c, values_only=True):
            row_dict, has_value = {}, False
            for i, val in enumerate(row):
                if val is not None:
                    has_value = True
                    if isinstance(val, (date, datetime)): val = val.isoformat()
                row_dict[headers[i]] = val if val is not None else ""
            if has_value: data.append(row_dict)
        return data

    def extract(self, excel_path: Path):
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        full_map = { "sheets": {} }
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            max_r, max_c = ws.max_row, ws.max_column
            merged_map = {}
            for m_range in ws.merged_cells.ranges:
                for r, c in m_range.cells:
                    if r == m_range.min_row and c == m_range.min_col:
                        merged_map[(r, c)] = {"colspan": m_range.max_col - m_range.min_col + 1, "rowspan": m_range.max_row - m_range.min_row + 1}
                    else: merged_map[(r, c)] = "skip"
            import html
            html_rows = ["<table border='1'>"]
            for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
                row_html = ["  <tr>"]
                for cell in row:
                    r, c = cell.row, cell.column
                    if merged_map.get((r, c)) == "skip": continue
                    val = cell.value
                    if isinstance(val, (date, datetime)): val = val.isoformat()
                    val_str = str(val) if val is not None else ""
                    attrs, spans = [], merged_map.get((r, c))
                    if isinstance(spans, dict):
                        if spans["colspan"] > 1: attrs.append(f'colspan="{spans["colspan"]}"')
                        if spans["rowspan"] > 1: attrs.append(f'rowspan="{spans["rowspan"]}"')
                    styles = []
                    if cell.fill and hasattr(cell.fill, "start_color") and cell.fill.start_color and cell.fill.start_color.index != '00000000':
                        c_idx = str(cell.fill.start_color.index)
                        if len(c_idx) == 8: c_idx = c_idx[2:] 
                        styles.append(f"background-color: #{c_idx}")
                    if styles: attrs.append(f'style="{"; ".join(styles)}"')
                    attr_str = " " + " ".join(attrs) if attrs else ""
                    safe_val = html.escape(val_str).replace('\n', '<br>')
                    row_html.append(f"<td{attr_str}>{safe_val}</td>")
                row_html.append("  </tr>")
                html_rows.append("".join(row_html))
            html_rows.append("</table>")
            full_map["sheets"][sheet_name] = {"html": "\n".join(html_rows), "media": self._extract_media(ws, sheet_name), "is_simple": self.is_simple_table(ws)}
            if full_map["sheets"][sheet_name]["is_simple"]: full_map["sheets"][sheet_name]["structured_data"] = self.extract_simple_table(ws)
        return full_map
