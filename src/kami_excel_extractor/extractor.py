import openpyxl
import html
from pathlib import Path
from openpyxl.utils import get_column_letter
import io
import logging
from typing import List, Dict, Any
from datetime import date, datetime
from PIL import Image, UnidentifiedImageError
from .utils import secure_filename, clean_kami_text

logger = logging.getLogger(__name__)

# セキュリティ設定: 画像の最大ピクセル数と最大バイト数
# デコンプレッションボム（DoS攻撃）対策
MAX_IMAGE_PIXELS = 25000000  # 25MP
MAX_IMAGE_BYTES = 20971520   # 20MB (20 * 1024 * 1024)

# Pillowのデフォルト制限を設定（Noneに設定されている場合の対策）
Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS

class MetadataExtractor:
    """Excelからメタデータとメディアを抽出するクラス"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)
        self.media_dir = self.output_dir / "media"
        self.media_dir.mkdir(parents=True, exist_ok=True)

    def _get_border_info(self, cell):
        """セルの罫線情報を取得する"""
        borders = {}
        if cell.border:
            # 線の太さやスタイルを取得
            if cell.border.left and cell.border.left.style: 
                borders["left"] = cell.border.left.style
            if cell.border.right and cell.border.right.style: 
                borders["right"] = cell.border.right.style
            if cell.border.top and cell.border.top.style: 
                borders["top"] = cell.border.top.style
            if cell.border.bottom and cell.border.bottom.style: 
                borders["bottom"] = cell.border.bottom.style
        return borders

    def _get_cell_style_string(self, cell):
        """セルのスタイル情報をCSS文字列として取得する"""
        styles = []
        
        # 背景色
        if cell.fill and hasattr(cell.fill, "start_color") and cell.fill.start_color:
            c_idx = str(cell.fill.start_color.index)
            # '00000000' はデフォルト（透明/自動）
            if c_idx not in ('00000000', '0'):
                if len(c_idx) == 8: # ARGB
                    c_idx = c_idx[2:]
                # 有効な16進数かチェック
                if all(c in "0123456789ABCDEFabcdef" for c in c_idx):
                    styles.append(f"background-color: #{c_idx}")

        # 罫線
        border_info = self._get_border_info(cell)
        for side, style in border_info.items():
            # thin, medium, thick などを CSS の border-style に近似
            width = "1px"
            if style in ('medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed'):
                width = "2px"
            elif style == 'thick':
                width = "3px"
            
            line_style = "solid"
            if 'dotted' in style.lower(): line_style = "dotted"
            elif 'dashed' in style.lower(): line_style = "dashed"
            
            styles.append(f"border-{side}: {width} {line_style} black")

        # フォント
        if cell.font:
            if cell.font.b: styles.append("font-weight: bold")
            if cell.font.i: styles.append("font-style: italic")

        return "; ".join(styles)

    def _extract_media(self, ws, sheet_name):
        media_info = []
        if not hasattr(ws, "_images"):
            return media_info

        for idx, img in enumerate(ws._images):
            # アンカー情報の取得 (OneCellAnchor, TwoCellAnchor, または文字列)
            row, col = None, None
            if hasattr(img.anchor, "_from"):
                row = img.anchor._from.row + 1
                col = img.anchor._from.col + 1
            elif isinstance(img.anchor, str):
                # "A1" 形式の文字列アンカーをパース
                from openpyxl.utils import coordinate_to_tuple
                try:
                    row, col = coordinate_to_tuple(img.anchor)
                except Exception:
                    pass
            
            if row is None or col is None:
                logger.warning(f"Could not determine anchor for image {idx} on sheet {sheet_name}")
                coord = "unknown"
            else:
                coord = f"{get_column_letter(col)}{row}"

            safe_sheet_name = secure_filename(sheet_name)
            image_filename = f"{safe_sheet_name}_img_{coord}_{idx}.png"
            save_path = self.media_dir / image_filename
            try:
                # 生のバイナリを取得
                raw_data = img.ref.read() if hasattr(img.ref, "read") else img.ref.getvalue()

                # セキュリティチェック: ファイルサイズが大きすぎる場合はスキップ
                if len(raw_data) > MAX_IMAGE_BYTES:
                    logger.warning(f"Skipping large image at {coord} on sheet {sheet_name} (size: {len(raw_data)} bytes)")
                    continue

                # Pillowを使って画像として読み込み、PNGとして再保存
                with Image.open(io.BytesIO(raw_data)) as pillow_img:
                    if pillow_img.mode in ("RGBA", "P"):
                        pillow_img = pillow_img.convert("RGB")
                    pillow_img.save(save_path, "PNG")

                media_info.append({"coord": coord, "filename": str(image_filename), "type": "image"})
            except (UnidentifiedImageError, OSError, ValueError, AttributeError, Image.DecompressionBombError) as e:
                # 🖼️ 改善: 画像自体の読み込みに失敗しても、座標情報だけは残す
                logger.warning(f"Failed to identify image at {coord} on sheet {sheet_name}: {e}. Keeping coordinate context.")
                media_info.append({"coord": coord, "filename": None, "type": "image", "error": "unidentified_format"})
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

    def _get_merged_cells_map(self, ws):
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

    def _cell_to_html_td(self, cell, span_info):
        """Convert a single cell to an HTML td element with rich styles and cleaned text."""
        val = cell.value
        if val is None:
            val_str = ""
        elif isinstance(val, (date, datetime)):
            val_str = val.isoformat()
        else:
            # セマンティック・クリーニングを適用
            val_str = str(clean_kami_text(val))

        attrs = []
        if isinstance(span_info, dict):
            if span_info["colspan"] > 1:
                attrs.append(f'colspan="{span_info["colspan"]}"')
            if span_info["rowspan"] > 1:
                attrs.append(f'rowspan="{span_info["rowspan"]}"')

        style_str = self._get_cell_style_string(cell)
        if style_str:
            attrs.append(f'style="{style_str}"')
        
        attrs.append(f'data-coord="{cell.coordinate}"')

        attr_str = " " + " ".join(attrs) if attrs else ""
        safe_val = html.escape(val_str).replace('\n', '<br>') if val_str else ""

        return f"<td{attr_str}>{safe_val}</td>"

    def _generate_cell_metadata(self, ws):
        merged_map = self._get_merged_cells_map(ws)
        cell_data = []
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                span = merged_map.get((cell.row, cell.column))
                if span == "skip": continue
                cell_info = {
                    "coord": cell.coordinate,
                    "row": cell.row,
                    "col": cell.column,
                    "value": str(clean_kami_text(cell.value)) if cell.value is not None else None,
                    "style": {
                        "borders": self._get_border_info(cell),
                        "bold": bool(cell.font.b) if cell.font else False,
                        "bg_color": str(cell.fill.start_color.index) if cell.fill and cell.fill.start_color else None
                    }
                }
                if isinstance(span, dict):
                    cell_info.update(span)
                cell_data.append(cell_info)
        return cell_data

    def _generate_html_table(self, ws):
        merged_map = self._get_merged_cells_map(ws)
        max_r, max_c = ws.max_row, ws.max_column
        html_rows = ["<table border='1' style='border-collapse: collapse; min-width: 100%;'>"]
        for row in ws.iter_rows(min_row=1, max_row=max_r, min_col=1, max_col=max_c):
            row_html = ["  <tr>"]
            for cell in row:
                r, c = cell.row, cell.column
                span_info = merged_map.get((r, c))
                if span_info == "skip": continue
                row_html.append(self._cell_to_html_td(cell, span_info))
            row_html.append("  </tr>")
            html_rows.append("".join(row_html))
        html_rows.append("</table>")
        return "\n".join(html_rows)

    def extract(self, excel_path: Path):
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        full_map = { "sheets": {} }
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            media_info = self._extract_media(ws, sheet_name)
            
            # 座標(coord)ごとにメディアをグループ化
            media_map = {}
            for m in media_info:
                coord = m.get("coord", "unknown")
                if coord not in media_map:
                    media_map[coord] = []
                media_map[coord].append(m)

            full_map["sheets"][sheet_name] = {
                "html": self._generate_html_table(ws),
                "cells": self._generate_cell_metadata(ws),
                "media": media_info,
                "media_map": media_map, # 座標ベースのマッピングを追加
                "is_simple": self.is_simple_table(ws)
            }
            if full_map["sheets"][sheet_name]["is_simple"]:
                full_map["sheets"][sheet_name]["structured_data"] = self.extract_simple_table(ws)
        return full_map
