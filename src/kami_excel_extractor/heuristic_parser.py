"""
LLMに依存しない、ルールベース/ヒューリスティックによるExcelデータ抽出エンジン。
"""

import logging
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple, Union
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from .utils import clean_kami_text

logger = logging.getLogger(__name__)


class HeuristicExcelParser:
    """
    結合セル、複数テーブル、キー・バリュー形式などの複雑なレイアウト（神エクセル）を、
    LLMを使用せずにルールベースで解析し、構造化データ（JSON）に変換するパーサー。
    """

    def __init__(self):
        pass

    def _get_merged_range_of_cell(self, ws: Worksheet, row: int, col: int) -> Optional[Any]:
        """指定された座標が含まれる結合セル範囲を返す。"""
        for r in ws.merged_cells.ranges:
            if r.min_row <= row <= r.max_row and r.min_col <= col <= r.max_col:
                return r
        return None

    def _get_cell_value(self, ws: Worksheet, row: int, col: int) -> Any:
        """結合セルを考慮してセルの値を取得する（結合セルの場合は左上セルの値を返す）。"""
        r_range = self._get_merged_range_of_cell(ws, row, col)
        if r_range:
            val = ws.cell(row=r_range.min_row, column=r_range.min_col).value
        else:
            val = ws.cell(row=row, column=col).value

        if isinstance(val, (date, datetime)):
            return val.isoformat()
        if isinstance(val, str):
            return clean_kami_text(val)
        return val

    def _is_row_empty(self, ws: Worksheet, row: int, min_col: int, max_col: int) -> bool:
        """指定された行の全セルが空（None または 空文字）かどうかを判定する。"""
        for col in range(min_col, max_col + 1):
            val = self._get_cell_value(ws, row, col)
            if val is not None and str(val).strip() != "":
                return False
        return True

    def _get_bounding_box(self, ws: Worksheet) -> Tuple[int, int, int, int]:
        """実質的にデータの存在するセルの範囲（最小行、最大行、最小列、最大列）を取得。"""
        min_r = ws.min_row or 1
        max_r = ws.max_row or 1
        min_c = ws.min_column or 1
        max_c = ws.max_column or 1

        # 実質的な最大行・列のトリミング
        real_max_r = min_r
        real_max_c = min_c

        for r in range(min_r, max_r + 1):
            row_empty = True
            for c in range(min_c, max_c + 1):
                val = ws.cell(row=r, column=c).value
                if val is not None and str(val).strip() != "":
                    row_empty = False
                    if c > real_max_c:
                        real_max_c = c
            if not row_empty:
                real_max_r = r

        return min_r, real_max_r, min_c, real_max_c

    def _split_into_segments(self, ws: Worksheet) -> List[Tuple[int, int]]:
        """空白行を区切りとして、シートを複数の論理セグメント（開始行、終了行のタプル）に分割する。"""
        min_r, max_r, min_c, max_c = self._get_bounding_box(ws)
        segments = []
        in_segment = False
        start_row = None

        for r in range(min_r, max_r + 1):
            is_empty = self._is_row_empty(ws, r, min_c, max_c)
            if not is_empty:
                if not in_segment:
                    in_segment = True
                    start_row = r
            else:
                if in_segment:
                    segments.append((start_row, r - 1))
                    in_segment = False

        if in_segment:
            segments.append((start_row, max_r))

        return segments

    def _is_key_value_block(self, ws: Worksheet, start_row: int, end_row: int, min_c: int, max_c: int) -> bool:
        """指定されたセグメントが、キー・バリュー形式（帳票）かどうかを判定するヒューリスティック。"""
        total_rows = end_row - start_row + 1
        if total_rows <= 1:
            return True

        colon_count = 0
        total_filled_cells = 0

        for r in range(start_row, end_row + 1):
            for c in range(min_c, max_c + 1):
                val = self._get_cell_value(ws, r, c)
                if val is not None and str(val).strip() != "":
                    total_filled_cells += 1
                    if ":" in str(val) or "：" in str(val):
                        colon_count += 1

        # コロンを含むセルがある場合
        if colon_count > 0:
            return True

        # データ密度が非常に低い、またはセルの数が極めて少ない場合は帳票/KV形式と判定
        if total_filled_cells <= 4:
            return True

        return False

    def _parse_table_segment(self, ws: Worksheet, start_row: int, end_row: int, min_c: int, max_c: int) -> List[Dict[str, Any]]:
        """セグメントを表（Table）形式としてパースする。"""
        # 多段ヘッダー対応: 1行目に結合セルがある場合、2行目もヘッダーの一部として扱う
        header_rows = 1
        has_merged_in_first_row = False
        for c in range(min_c, max_c + 1):
            r_range = self._get_merged_range_of_cell(ws, start_row, c)
            if r_range:
                has_merged_in_first_row = True
                header_rows = max(header_rows, r_range.max_row - start_row + 1)

        # 横方向のみの結合があり、かつ次の行が存在する場合、header_rowsを2にする
        if has_merged_in_first_row and header_rows == 1 and (start_row + 1 <= end_row):
            if not self._is_row_empty(ws, start_row + 1, min_c, max_c):
                header_rows = 2

        # ヘッダー情報の抽出と結合セルの名寄せ
        raw_headers = []
        for r_h in range(start_row, start_row + header_rows):
            row_headers = []
            for c in range(min_c, max_c + 1):
                val = self._get_cell_value(ws, r_h, c)
                val_str = str(val).strip() if val is not None else ""
                row_headers.append(val_str)
            raw_headers.append(row_headers)

        # 最終的な平坦化されたヘッダー名の作成 (例: "基本情報_氏名")
        final_headers = []
        for col_idx in range(len(raw_headers[0])):
            header_parts = []
            for r_idx in range(header_rows):
                part = raw_headers[r_idx][col_idx]
                if part and part not in header_parts:
                    header_parts.append(part)

            final_name = "_".join(header_parts) if header_parts else f"Column_{col_idx + 1}"
            final_headers.append(final_name)

        # データ行のパース
        data_list = []
        data_start_row = start_row + header_rows

        for r in range(data_start_row, end_row + 1):
            # 行全体が空の場合はスキップ
            if self._is_row_empty(ws, r, min_c, max_c):
                continue

            row_dict = {}
            row_has_value = False
            for i, col in enumerate(range(min_c, max_c + 1)):
                header_name = final_headers[i]
                val = self._get_cell_value(ws, r, col)
                if val is not None and str(val).strip() != "":
                    row_dict[header_name] = val
                    row_has_value = True

            if row_has_value:
                data_list.append(row_dict)

        return data_list

    def _parse_key_value_segment(self, ws: Worksheet, start_row: int, end_row: int, min_c: int, max_c: int) -> Dict[str, Any]:
        """セグメントをキー・バリュー（フォーム）形式としてパースする。"""
        kv_data = {}
        processed_cells = set()

        for r in range(start_row, end_row + 1):
            for c in range(min_c, max_c + 1):
                cell_coord = (r, c)
                if cell_coord in processed_cells:
                    continue

                val = self._get_cell_value(ws, r, c)
                if val is None or str(val).strip() == "":
                    continue

                val_str = str(val).strip()

                # ラベル（キー）と判定される条件:
                # - 文字列であり、コロンで終わる、または隣のセルがデータセル
                is_label = isinstance(val, str) and (
                    val_str.endswith(":") or val_str.endswith("：") or len(val_str) <= 30
                )

                if is_label:
                    key_name = val_str.rstrip(":：")

                    # 結合セル範囲を取得
                    r_range = self._get_merged_range_of_cell(ws, r, c)
                    span_cols = (r_range.max_col - r_range.min_col + 1) if r_range else 1
                    span_rows = (r_range.max_row - r_range.min_row + 1) if r_range else 1

                    # ラベルに含まれる全セルを処理済みとしてマーク
                    for rs in range(r, r + span_rows):
                        for cs in range(c, c + span_cols):
                            processed_cells.add((rs, cs))

                    # バリューを探す
                    # 1. まず同じ行の右側のセルを探索
                    val_col = c + span_cols
                    found_value = False
                    if val_col <= max_c:
                        val_cell_val = self._get_cell_value(ws, r, val_col)
                        if val_cell_val is not None and str(val_cell_val).strip() != "":
                            kv_data[key_name] = val_cell_val
                            found_value = True

                            # バリューの結合セル範囲もマーク
                            v_range = self._get_merged_range_of_cell(ws, r, val_col)
                            v_span_cols = (v_range.max_col - v_range.min_col + 1) if v_range else 1
                            v_span_rows = (v_range.max_row - v_range.min_row + 1) if v_range else 1
                            for rs in range(r, r + v_span_rows):
                                for cs in range(val_col, val_col + v_span_cols):
                                    processed_cells.add((rs, cs))

                    # 2. 右側になければ直下のセルを探索
                    if not found_value:
                        val_row = r + span_rows
                        if val_row <= end_row:
                            val_cell_val = self._get_cell_value(ws, val_row, c)
                            if val_cell_val is not None and str(val_cell_val).strip() != "":
                                kv_data[key_name] = val_cell_val
                                found_value = True

                                v_range = self._get_merged_range_of_cell(ws, val_row, c)
                                v_span_cols = (v_range.max_col - v_range.min_col + 1) if v_range else 1
                                v_span_rows = (v_range.max_row - v_range.min_row + 1) if v_range else 1
                                for rs in range(val_row, val_row + v_span_rows):
                                    for cs in range(c, c + v_span_cols):
                                        processed_cells.add((rs, cs))

                    # バリューが見つからなかったがキーが存在する場合は空文字をセット
                    if not found_value and key_name:
                        kv_data[key_name] = ""
                else:
                    # ラベルでない単独の値セル
                    kv_data[f"Value_{get_column_letter(c)}{r}"] = val
                    processed_cells.add((r, c))

        return kv_data

    def parse_worksheet(self, ws: Worksheet) -> Dict[str, Any]:
        """ワークシートをパースして構造化された辞書を返す。"""
        # 単純な表（結合なし、見出し＋データ）の場合は高速な簡易パスを使用
        if not ws.merged_cells.ranges and ws.max_row >= 2 and ws.max_column >= 1:
            first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
            header_count = sum(1 for v in first_row if v is not None)
            if header_count >= 2:
                # 簡易表抽出
                rows_gen = ws.iter_rows(values_only=True)
                try:
                    header_row = next(rows_gen)
                    headers = [str(v or f"Column{i + 1}") for i, v in enumerate(header_row)]
                    data = []
                    for row in rows_gen:
                        row_dict = {}
                        for i, v in enumerate(row):
                            if v is not None:
                                row_dict[headers[i]] = v.isoformat() if isinstance(v, (date, datetime)) else v
                        if row_dict:
                            data.append(row_dict)
                    return {"data": data}
                except StopIteration:
                    pass

        # 複雑なテーブル/フォームのヒューリスティック解析
        min_r, max_r, min_c, max_c = self._get_bounding_box(ws)
        segments = self._split_into_segments(ws)

        if not segments:
            return {"data": []}

        # セグメントが1つだけで表形式の場合
        if len(segments) == 1:
            start_row, end_row = segments[0]
            if not self._is_key_value_block(ws, start_row, end_row, min_c, max_c):
                table_data = self._parse_table_segment(ws, start_row, end_row, min_c, max_c)
                return {"data": table_data}
            else:
                kv_data = self._parse_key_value_segment(ws, start_row, end_row, min_c, max_c)
                return {"data": kv_data}

        # 複数セグメントがある場合
        combined_result = {}
        table_count = 0
        form_count = 0

        for idx, (start_row, end_row) in enumerate(segments):
            if not self._is_key_value_block(ws, start_row, end_row, min_c, max_c):
                table_data = self._parse_table_segment(ws, start_row, end_row, min_c, max_c)
                if table_data:
                    table_count += 1
                    combined_result[f"table_{table_count}"] = table_data
            else:
                kv_data = self._parse_key_value_segment(ws, start_row, end_row, min_c, max_c)
                if kv_data:
                    form_count += 1
                    combined_result[f"form_{form_count}"] = kv_data

        # 1つのテーブル、またはすべてのセグメントがフォームの場合、結果を平坦化する
        if len(combined_result) == 1:
            key = list(combined_result.keys())[0]
            return {"data": combined_result[key]}

        # すべてがフォームの場合、平坦な辞書としてマージする
        if all(k.startswith("form_") for k in combined_result.keys()):
            merged_kv = {}
            for kv in combined_result.values():
                merged_kv.update(kv)
            return {"data": merged_kv}

        return {"data": combined_result}
