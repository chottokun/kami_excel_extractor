from pathlib import Path
from unittest.mock import MagicMock

import pytest

from kami_excel_extractor.extractor import MetadataExtractor


def test_get_cell_style_string(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Case 1: 8-digit color index (e.g., AARRGGBB) -> Should strip first 2 digits
    cell_8 = MagicMock()
    cell_8.fill.start_color.index = "FFFF0000"  # Red
    cell_8.border = None
    cell_8.font = None
    style_8 = extractor._get_cell_style_string(cell_8)
    assert "background-color: #FF0000" in style_8

    # Case 2: 6-digit color index (e.g., RRGGBB) -> Should keep as is
    cell_6 = MagicMock()
    cell_6.fill.start_color.index = "00FF00"  # Green
    cell_6.border = None
    cell_6.font = None
    style_6 = extractor._get_cell_style_string(cell_6)
    assert "background-color: #00FF00" in style_6

    # Case 3: '00000000' color index -> Should result in no background style
    cell_0 = MagicMock()
    cell_0.fill.start_color.index = "00000000"
    cell_0.border = None
    cell_0.font = None
    style_0 = extractor._get_cell_style_string(cell_0)
    assert "background-color" not in style_0


def test_cell_to_html_td_styles(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    cell = MagicMock()
    cell.value = "TestValue"
    cell.coordinate = "A1"

    style_str = "background-color: #FF0000"
    unit = "JPY"

    html_out = extractor._cell_to_html_td(cell, None, style_str, unit)

    assert 'style="background-color: #FF0000"' in html_out
    assert 'data-unit="JPY"' in html_out
    assert ">TestValue</td>" in html_out
