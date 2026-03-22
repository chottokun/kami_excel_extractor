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
            
            merged_map = {}
            for m_range in ws.merged_cells.ranges:
                for r, c in m_range.cells:
                    if r == m_range.min_row and c == m_range.min_col:
                        merged_map[(r, c)] = {
                            "colspan": m_range.max_col - m_range.min_col + 1,
                            "rowspan": m_range.max_row - m_range.min_row + 1
                        }
                    else:
                        merged_map[(r, c)] = "skip"
            
            import html
            html_rows = ["<table border='1'>"]
            
            for r in range(1, ws.max_row + 1):
                row_html = ["  <tr>"]
                for c in range(1, ws.max_column + 1):
                    if merged_map.get((r, c)) == "skip":
                        continue
                    
                    cell = ws.cell(row=r, column=c)
                    val = str(cell.value) if cell.value is not None else ""
                    
                    attrs = []
                    spans = merged_map.get((r, c))
                    if isinstance(spans, dict):
                        if spans["colspan"] > 1: attrs.append(f'colspan="{spans["colspan"]}"')
                        if spans["rowspan"] > 1: attrs.append(f'rowspan="{spans["rowspan"]}"')
                        
                    styles = []
                    if cell.fill and hasattr(cell.fill, "start_color") and cell.fill.start_color and cell.fill.start_color.index != '00000000':
                        c_idx = str(cell.fill.start_color.index)
                        if len(c_idx) == 8: c_idx = c_idx[2:] # ARGB to RGB roughly
                        styles.append(f"background-color: #{c_idx}")
                        
                    if styles:
                        attrs.append(f'style="{"; ".join(styles)}"')
                        
                    attr_str = " " + " ".join(attrs) if attrs else ""
                    safe_val = html.escape(val).replace('\n', '<br>')
                    row_html.append(f"<td{attr_str}>{safe_val}</td>")
                row_html.append("  </tr>")
                html_rows.append("".join(row_html))
            html_rows.append("</table>")
            
            full_map["sheets"][sheet_name] = {
                "html": "\n".join(html_rows),
                "media": self._extract_media(ws, sheet_name)
            }
        return full_map
