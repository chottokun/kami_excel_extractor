"""
Excelファイルから詳細なメタデータ、視覚スタイル、およびロジックを抽出するエンジン。
"""

import logging
import io
import html
import openpyxl
from pathlib import Path
from openpyxl.utils import get_column_letter, coordinate_to_tuple
from typing import List, Dict, Any, Optional, Union, Tuple
from datetime import date, datetime
from PIL import Image, UnidentifiedImageError
from .utils import secure_filename, clean_kami_text

logger = logging.getLogger(__name__)

# セキュリティ設定: 画像の最大ピクセル数と最大バイト数
MAX_IMAGE_PIXELS = 25000000  # 25MP
MAX_IMAGE_BYTES = 20971520   # 20MB
Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS

class MetadataExtractor:
    """
    Excelワークブックからテキスト、構造、スタイル、および埋め込みメディアを抽出する高精度エクストラクター。
    
    Attributes:
        output_dir (Path): 抽出されたメディア（画像）を保存するディレクトリ。
        media_dir (Path): メディアファイルの具体的な保存先。
    """
    
    def __init__(self, output_dir: Union[str, Path]):
        """
        MetadataExtractorを初期化する。
        
        Args:
            output_dir: 解析結果（主に画像）を出力するディレクトリパス。
        """
        self.output_dir = Path(output_dir)
        self.media_dir = self.output_dir / "media"
        self.media_dir.mkdir(parents=True, exist_ok=True)

    def _get_border_info(self, cell: openpyxl.cell.Cell) -> Dict[str, str]:
        """
        セルの罫線情報を取得し、各辺のスタイルを辞書形式で返す。
        
        Args:
            cell: openpyxlのセルオブジェクト。
            
        Returns:
            Dict[str, str]: {'left': 'thin', 'top': 'thick', ...} 形式の辞書。
        """
        borders = {}
        if cell.border:
            sides = ['left', 'right', 'top', 'bottom']
            for side in sides:
                border_side = getattr(cell.border, side)
                if border_side and border_side.style:
                    borders[side] = border_side.style
        return borders

    def _get_cell_style_string(self, cell: openpyxl.cell.Cell) -> str:
        """
        セルの視覚的属性（色、線、フォント）をCSS形式の文字列に変換する。
        
        LLMがテーブルの構造（ヘッダー、セクションの区切り）を理解するための重要なヒントとなる。
        
        Args:
            cell: openpyxlのセルオブジェクト。
            
        Returns:
            str: "background-color: #FFFFFF; border-top: 1px solid black;" 等のCSS文字列。
        """
        styles = []
        
        # 背景色の抽出 (ARGBをRGBに変換)
        if cell.fill and hasattr(cell.fill, "start_color") and cell.fill.start_color:
            c_idx = str(cell.fill.start_color.index)
            if c_idx not in ('00000000', '0'):
                if len(c_idx) == 8: # ARGB (Alpha-RGB)
                    c_idx = c_idx[2:]
                if all(c in "0123456789ABCDEFabcdef" for c in c_idx):
                    styles.append(f"background-color: #{c_idx}")

        # 罫線情報をCSSのborder属性に近似
        border_info = self._get_border_info(cell)
        style_map = {
            'medium': '2px solid',
            'thick': '3px solid',
            'thin': '1px solid',
            'dashed': '1px dashed',
            'dotted': '1px dotted'
        }
        
        for side, style in border_info.items():
            css_style = style_map.get(style, '1px solid')
            styles.append(f"border-{side}: {css_style} black")

        # フォントウェイト
        if cell.font:
            if cell.font.b: styles.append("font-weight: bold")
            if cell.font.i: styles.append("font-style: italic")

        return "; ".join(styles)

    def _get_unit_info(self, cell: openpyxl.cell.Cell) -> Optional[str]:
        """
        セルの表示形式(Number Format)からデータの単位や型を推測する。
        
        Args:
            cell: openpyxlのセルオブジェクト。
            
        Returns:
            Optional[str]: 'JPY', 'PERCENT', 'DATE' 等の識別子。
        """
        fmt = cell.number_format
        if not fmt or fmt == 'General':
            return None
        
        fmt_lower = fmt.lower()
        if '¥' in fmt or 'jpy' in fmt_lower: return 'JPY'
        if '$' in fmt: return 'USD'
        if '%' in fmt: return 'PERCENT'
        if 'yy' in fmt_lower or 'mm' in fmt_lower or 'dd' in fmt_lower: return 'DATE'
        return fmt

    def _extract_media(self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str) -> List[Dict[str, Any]]:
        """
        ワークシートから埋め込み画像（図、グラフ、写真）を抽出し保存する。
        """
        media_info = []
        if not hasattr(ws, "_images"):
            return media_info

        for idx, img in enumerate(ws._images):
            row, col = None, None
            anchor = img.anchor
            
            # 各種アンカー形式をパース
            if hasattr(anchor, "_from"): # TwoCellAnchor
                row = anchor._from.row + 1
                col = anchor._from.col + 1
            elif hasattr(anchor, "row"): # OneCellAnchor
                row = anchor.row + 1
                col = anchor.col + 1
            elif isinstance(anchor, str): # String anchor (e.g. "A1")
                try:
                    row, col = coordinate_to_tuple(anchor)
                except Exception:
                    pass
            
            coord = f"{get_column_letter(col)}{row}" if (row is not None and col is not None) else "unknown"
            safe_sheet_name = secure_filename(sheet_name)
            image_filename = f"{safe_sheet_name}_img_{coord}_{idx}.png"
            save_path = self.media_dir / image_filename
            
            item = {"coord": coord, "filename": str(image_filename), "type": "image"}
            
            try:
                # 🔒 Security Fix: Read stream in chunks to prevent DoS via memory exhaustion
                if hasattr(img.ref, "read"):
                    raw_data_buf = io.BytesIO()
                    total_read = 0
                    chunk_size = 8192
                    while True:
                        chunk = img.ref.read(chunk_size)
                        if not chunk:
                            break
                        total_read += len(chunk)
                        if total_read > MAX_IMAGE_BYTES:
                            logger.warning(f"Skipping large image at {coord} on {sheet_name} (stream exceeds limit)")
                            break
                        raw_data_buf.write(chunk)

                    if total_read > MAX_IMAGE_BYTES:
                        continue
                    raw_data = raw_data_buf.getvalue()
                else:
                    raw_data = img.ref.getvalue()
                    if len(raw_data) > MAX_IMAGE_BYTES:
                        logger.warning(f"Skipping large image at {coord} on {sheet_name}")
                        continue

                with Image.open(io.BytesIO(raw_data)) as pillow_img:
                    if pillow_img.mode in ("RGBA", "P"):
                        pillow_img = pillow_img.convert("RGB")
                    pillow_img.save(save_path, "PNG")
                media_info.append(item)
            except Exception as e:
                logger.warning(f"Failed to extract image at {coord} on sheet {sheet_name}: {e}")
                item["filename"] = None
                item["error"] = "unidentified_format"
                media_info.append(item)
                
        return media_info

    def _get_merged_cells_map(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> Dict[Tuple[int, int], Union[str, Dict[str, int]]]:
        """
        シート内の結合セル情報をマップ化する。
        
        Returns:
            Dict: 左上セルの座標(r, c)をキーとし、スパン情報を値に持つ辞書。
        """
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
        return merged_map

    def _cell_to_html_td(self, cell: openpyxl.cell.Cell, span_info: Union[str, Dict], formula: Optional[str] = None) -> str:
        """
        単一のセルを、詳細属性付きのHTML <td> タグに変換する。
        """
        val = cell.value
        val_str = val.isoformat() if isinstance(val, (date, datetime)) else str(clean_kami_text(val)) if val is not None else ""

        attrs = [f'data-coord="{cell.coordinate}"']
        if isinstance(span_info, dict):
            if span_info.get("colspan", 1) > 1: attrs.append(f'colspan="{span_info["colspan"]}"')
            if span_info.get("rowspan", 1) > 1: attrs.append(f'rowspan="{span_info["rowspan"]}"')

        style_str = self._get_cell_style_string(cell)
        if style_str: attrs.append(f'style="{style_str}"')
        
        if formula and str(formula).startswith('='):
            attrs.append(f'data-formula="{html.escape(str(formula))}"')
        
        unit = self._get_unit_info(cell)
        if unit: attrs.append(f'data-unit="{html.escape(unit)}"')

        attr_str = " " + " ".join(attrs) if attrs else ""
        safe_val = html.escape(val_str).replace('\n', '<br>')
        return f"<td{attr_str}>{safe_val}</td>"

    def is_simple_table(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> bool:
        """シートが単純な表形式（結合なし、1行目見出し）かどうかを判定する。"""
        if ws.merged_cells.ranges: return False
        if ws.max_row < 2 or ws.max_column < 1: return False

        # 1行目に少なくとも2つの非空セルがあるか確認（早めに終了）
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
        header_count = 0
        for val in first_row:
            if val is not None:
                header_count += 1
                if header_count >= 2:
                    return True
        return False

    def _extract_row_dict(self, row: Tuple[Any, ...], headers: List[str]) -> Dict[str, Any]:
        """1行のデータを辞書形式に変換する（日付処理・None除外を含む）。"""
        row_dict = {}
        for i, v in enumerate(row):
            if v is not None:
                row_dict[headers[i]] = v.isoformat() if isinstance(v, (date, datetime)) else v
        return row_dict

    def extract_simple_table(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> List[Dict[str, Any]]:
        """単純な表形式のシートからデータを高速に抽出する。"""
        rows_gen = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows_gen)
        except StopIteration:
            return []

        headers = [str(v or f"Column{i+1}") for i, v in enumerate(header_row)]
        data = []
        for row in rows_gen:
            row_dict = self._extract_row_dict(row, headers)
            if row_dict:
                data.append(row_dict)
        return data

    def _get_bounding_box(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> Tuple[int, int, int, int]:
        """
        データまたは書式が存在する実質的な範囲（最小行、最大行、最小列、最大列）を特定する。
        """
        min_r, max_r = 1, 0
        min_c, max_c = 1, 0

        # データがあるセルの範囲を取得
        if ws.max_row > 0:
            for r in range(ws.max_row, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for c in range(1, ws.max_column + 1)):
                    max_r = r
                    break
            for c in range(ws.max_column, 0, -1):
                if any(ws.cell(row=r, column=c).value is not None for r in range(1, ws.max_row + 1)):
                    max_c = c
                    break
        
        # 結合セルや画像がある範囲も含める
        for m_range in ws.merged_cells.ranges:
            max_r = max(max_r, m_range.max_row)
            max_c = max(max_c, m_range.max_col)
        
        if hasattr(ws, "_images"):
            for img in ws._images:
                if hasattr(img.anchor, "_from"):
                    max_r = max(max_r, img.anchor._from.row + 1)
                    max_c = max(max_c, img.anchor._from.col + 1)

        return 1, max_r, 1, max_c

    def _generate_metadata_and_html(self, ws: openpyxl.worksheet.worksheet.Worksheet, ws_formula: Optional[openpyxl.worksheet.worksheet.Worksheet] = None, merged_map: Optional[Dict] = None) -> Tuple[str, List[Dict[str, Any]]]:
        """詳細メタデータとHTMLテーブルを同時に生成する。"""
        if merged_map is None:
            merged_map = self._get_merged_cells_map(ws)

        min_r, max_r, min_c, max_c = self._get_bounding_box(ws)
        cell_metadata = []
        html_rows = ["<table border='1' style=\"border-collapse: collapse; min-width: 100%;\">"]

        for r in range(min_r, max_r + 1):
            row_html = ["  <tr>"]
            row_has_data = False
            
            current_row_html = []
            for c in range(min_c, max_c + 1):
                cell = ws.cell(row=r, column=c)
                span = merged_map.get((r, c))
                if span == "skip": continue

                formula = ws_formula.cell(row=r, column=c).value if ws_formula else None
                if cell.value is not None or formula is not None:
                    row_has_data = True

                # メタデータの構築
                cell_info = {
                    "coord": cell.coordinate, "row": r, "col": c,
                    "value": str(clean_kami_text(cell.value)) if cell.value is not None else None,
                    "formula": formula if str(formula).startswith('=') else None,
                    "unit": self._get_unit_info(cell),
                    "style": {"borders": self._get_border_info(cell), "bold": bool(cell.font.b if cell.font else False)}
                }
                if isinstance(span, dict): cell_info.update(span)
                cell_metadata.append(cell_info)

                # HTMLテーブル行の構築
                td_html = self._cell_to_html_td(cell, span, formula=formula)
                current_row_html.append(td_html)
            
            # データがない空行は、先行するデータがある場合のみ出力するなどの調整も可能だが、
            # ここでは bounding box 内の全行を出す（スタイルがある可能性があるため）。
            row_html.extend(current_row_html)
            row_html.append("  </tr>")
            html_rows.append("".join(row_html))
        
        html_rows.append("</table>")
        return "\n".join(html_rows), cell_metadata

    def _generate_cell_metadata(self, ws: openpyxl.worksheet.worksheet.Worksheet, ws_formula: Optional[openpyxl.worksheet.worksheet.Worksheet] = None, merged_map: Optional[Dict] = None) -> List[Dict[str, Any]]:
        """詳細メタデータを生成する。"""
        _, cell_metadata = self._generate_metadata_and_html(ws, ws_formula=ws_formula, merged_map=merged_map)
        return cell_metadata

    def _generate_html_table(self, ws: openpyxl.worksheet.worksheet.Worksheet, ws_formula: Optional[openpyxl.worksheet.worksheet.Worksheet] = None, merged_map: Optional[Dict] = None) -> str:
        """シートからHTMLテーブルを生成する。"""
        html_table, _ = self._generate_metadata_and_html(ws, ws_formula=ws_formula, merged_map=merged_map)
        return html_table

    def extract(self, excel_path: Path, include_logic: bool = False) -> Dict[str, Any]:
        """
        Excelファイルを解析し、詳細な構造、スタイル、ロジック、メディアを抽出する。
        
        Args:
            excel_path: 解析対象のエクセルファイルのパス。
            include_logic: 計算式(formula)の抽出を有効にするかどうか。
            
        Returns:
            Dict: 全シートの解析データを含む辞書。
        """
        # 値の抽出用に data_only=True でロード
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        wb_formula = openpyxl.load_workbook(excel_path, data_only=False) if include_logic else None

        full_map = { "sheets": {} }
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws_f = wb_formula[sheet_name] if wb_formula else None
            
            media_info = self._extract_media(ws, sheet_name)
            media_map = {}
            for m in media_info:
                coord = m.get("coord", "unknown")
                media_map.setdefault(coord, []).append(m)

            # 詳細メタデータの生成
            merged_map = self._get_merged_cells_map(ws)
            html_table, cell_metadata = self._generate_metadata_and_html(ws, ws_formula=ws_f, merged_map=merged_map)

            full_map["sheets"][sheet_name] = {
                "html": html_table,
                "cells": cell_metadata,
                "media": media_info,
                "media_map": media_map,
                "is_simple": self.is_simple_table(ws)
            }
            if full_map["sheets"][sheet_name]["is_simple"]:
                full_map["sheets"][sheet_name]["structured_data"] = self.extract_simple_table(ws)
                
        return full_map
