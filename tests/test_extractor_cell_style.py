from pathlib import Path
from unittest.mock import MagicMock

import pytest

from kami_excel_extractor.extractor import MetadataExtractor


def test_cell_to_html_td_styles(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Case 1: 8-digit color index (e.g., AARRGGBB) -> Should strip first 2 digits
    cell_8 = MagicMock()
    cell_8.value = "Test8"
    cell_8.coordinate = "A1"
    cell_8.fill.start_color.index = "FFFF0000"  # Red
    style_str_8 = extractor._get_cell_style_string(cell_8)
    unit_8 = extractor._get_unit_info(cell_8)
    html_8 = extractor._cell_to_html_td(cell_8, None, style_str_8, unit_8)
    assert "background-color: #FF0000" in html_8
    assert ">Test8</td>" in html_8

    # Case 2: 6-digit color index (e.g., RRGGBB) -> Should keep as is
    cell_6 = MagicMock()
    cell_6.value = "Test6"
    cell_6.coordinate = "A2"
    cell_6.fill.start_color.index = "00FF00"  # Green
    style_str_6 = extractor._get_cell_style_string(cell_6)
    unit_6 = extractor._get_unit_info(cell_6)
    html_6 = extractor._cell_to_html_td(cell_6, None, style_str_6, unit_6)
    assert "background-color: #00FF00" in html_6
    assert ">Test6</td>" in html_6

    # Case 3: '00000000' color index -> Should result in no background style
    cell_0 = MagicMock()
    cell_0.value = "Test0"
    cell_0.coordinate = "A3"
    cell_0.fill.start_color.index = "00000000"
    style_str_0 = extractor._get_cell_style_string(cell_0)
    unit_0 = extractor._get_unit_info(cell_0)
    html_0 = extractor._cell_to_html_td(cell_0, None, style_str_0, unit_0)
    assert "background-color" not in html_0
    assert ">Test0</td>" in html_0

    # Case 4: No fill
    cell_no_fill = MagicMock()
    cell_no_fill.value = "TestNoFill"
    cell_no_fill.coordinate = "A4"
    cell_no_fill.fill = None
    style_str_no_fill = extractor._get_cell_style_string(cell_no_fill)
    unit_no_fill = extractor._get_unit_info(cell_no_fill)
    html_no_fill = extractor._cell_to_html_td(cell_no_fill, None, style_str_no_fill, unit_no_fill)
    assert "background-color" not in html_no_fill
    assert ">TestNoFill</td>" in html_no_fill

    # Case 5: Fill without start_color
    cell_no_color = MagicMock()
    cell_no_color.value = "TestNoColor"
    cell_no_color.coordinate = "A5"
    cell_no_color.fill.start_color = None
    style_str_no_color = extractor._get_cell_style_string(cell_no_color)
    unit_no_color = extractor._get_unit_info(cell_no_color)
    html_no_color = extractor._cell_to_html_td(cell_no_color, None, style_str_no_color, unit_no_color)
    assert "background-color" not in html_no_color
    assert ">TestNoColor</td>" in html_no_color
