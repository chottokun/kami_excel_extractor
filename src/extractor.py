import openpyxl
import json
from openpyxl.utils import get_column_letter

def get_border_info(cell):
    borders = {}
    if cell.border:
        if cell.border.left.style: borders["L"] = cell.border.left.style
        if cell.border.right.style: borders["R"] = cell.border.right.style
        if cell.border.top.style: borders["T"] = cell.border.top.style
        if cell.border.bottom.style: borders["B"] = cell.border.bottom.style
    return borders

def extract_universal_map(filename):
    wb = openpyxl.load_workbook(filename, data_only=True)
    ws = wb.active
    
    merged_ranges = [str(r) for r in ws.merged_cells.ranges]
    
    cells_map = []
    # 有効なデータ範囲を特定
    max_row = ws.max_row
    max_col = ws.max_column
    
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            
            # 空セルかつスタイルなしはスキップして圧縮
            has_style = cell.fill and cell.fill.start_color.index != '00000000'
            has_border = bool(get_border_info(cell))
            
            if cell.value is None and not has_style and not has_border:
                continue
                
            cell_info = {
                "c": cell.coordinate, # coordinate
                "v": str(cell.value) if cell.value is not None else "", # value
            }
            
            # スタイル情報はヒントとして最小限に
            if has_style:
                cell_info["bg"] = str(cell.fill.start_color.index)
            if has_border:
                cell_info["b"] = get_border_info(cell)
            
            # 結合セルの判定
            for m_range in merged_ranges:
                if cell.coordinate in openpyxl.worksheet.cell_range.CellRange(m_range):
                    cell_info["m"] = m_range # merged range
                    break
            
            cells_map.append(cell_info)
            
    return cells_map

if __name__ == "__main__":
    test_map = extract_universal_map("sample_hoganshi.xlsx")
    with open("sample_map.json", "w", encoding="utf-8") as f:
        json.dump(test_map, f, ensure_ascii=False, indent=2)
    print(f"Extracted {len(test_map)} cells to sample_map.json")
