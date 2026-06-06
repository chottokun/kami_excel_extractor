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
    assert style_8 == "background-color: #FF0000"

    # Case 2: 6-digit color index (e.g., RRGGBB) -> Should keep as is
    cell_6 = MagicMock()
    cell_6.fill.start_color.index = "00FF00"  # Green
    cell_6.border = None
    cell_6.font = None
    style_6 = extractor._get_cell_style_string(cell_6)
    assert style_6 == "background-color: #00FF00"

    # Case 3: '00000000' color index -> Should result in no background style
    cell_0 = MagicMock()
    cell_0.fill.start_color.index = "00000000"
    cell_0.border = None
    cell_0.font = None
    style_0 = extractor._get_cell_style_string(cell_0)
    assert style_0 == ""

    # Case 4: No fill
    cell_no_fill = MagicMock()
    cell_no_fill.fill = None
    cell_no_fill.border = None
    cell_no_fill.font = None
    style_no_fill = extractor._get_cell_style_string(cell_no_fill)
    assert style_no_fill == ""

    # Case 5: Fill without start_color
    cell_no_color = MagicMock()
    cell_no_color.fill.start_color = None
    cell_no_color.border = None
    cell_no_color.font = None
    style_no_color = extractor._get_cell_style_string(cell_no_color)
    assert style_no_color == ""


def test_cell_to_html_td_basic(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.coordinate = "A1"

    # Test that it correctly assembles the HTML with provided values
    html = extractor._cell_to_html_td(cell, {"colspan": 2}, "Hello\nWorld", "color: red", "JPY")
    assert 'data-coord="A1"' in html
    assert 'colspan="2"' in html
    assert 'style="color: red"' in html
    assert 'data-unit="JPY"' in html
    assert 'Hello<br>World' in html
