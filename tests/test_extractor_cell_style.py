import pytest
from unittest.mock import MagicMock
from pathlib import Path
from kami_excel_extractor.extractor import MetadataExtractor

def test_cell_to_html_td_styles(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Case 1: 8-digit color index (e.g., AARRGGBB) -> Should strip first 2 digits
    cell_8 = MagicMock()
    cell_8.value = "Test8"
    cell_8.fill.start_color.index = "FFFF0000" # Red
    html_8 = extractor._cell_to_html_td(cell_8, None)
    assert 'style="background-color: #FF0000"' in html_8
    assert ">Test8</td>" in html_8

    # Case 2: 6-digit color index (e.g., RRGGBB) -> Should keep as is
    cell_6 = MagicMock()
    cell_6.value = "Test6"
    cell_6.fill.start_color.index = "00FF00" # Green
    html_6 = extractor._cell_to_html_td(cell_6, None)
    assert 'style="background-color: #00FF00"' in html_6
    assert ">Test6</td>" in html_6

    # Case 3: '00000000' color index -> Should result in no style
    cell_0 = MagicMock()
    cell_0.value = "Test0"
    cell_0.fill.start_color.index = "00000000"
    html_0 = extractor._cell_to_html_td(cell_0, None)
    assert 'style=' not in html_0
    assert ">Test0</td>" in html_0

    # Case 4: No fill
    cell_no_fill = MagicMock()
    cell_no_fill.value = "TestNoFill"
    cell_no_fill.fill = None
    html_no_fill = extractor._cell_to_html_td(cell_no_fill, None)
    assert 'style=' not in html_no_fill
    assert ">TestNoFill</td>" in html_no_fill

    # Case 5: Fill without start_color
    cell_no_color = MagicMock()
    cell_no_color.value = "TestNoColor"
    cell_no_color.fill.start_color = None
    html_no_color = extractor._cell_to_html_td(cell_no_color, None)
    assert 'style=' not in html_no_color
    assert ">TestNoColor</td>" in html_no_color
