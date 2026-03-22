from unittest.mock import MagicMock
from kami_excel_extractor.extractor import MetadataExtractor

def test_get_border_info_no_border(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.border = None

    assert extractor._get_border_info(cell) == {}

def test_get_border_info_all_borders(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    cell.border.left.style = "thin"
    cell.border.right.style = "thick"
    cell.border.top.style = "dashed"
    cell.border.bottom.style = "double"

    expected = {
        "L": "thin",
        "R": "thick",
        "T": "dashed",
        "B": "double"
    }
    assert extractor._get_border_info(cell) == expected

def test_get_border_info_partial_borders(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    cell.border.left.style = "thin"
    cell.border.right.style = None
    cell.border.top.style = "thick"
    cell.border.bottom.style = None

    expected = {
        "L": "thin",
        "T": "thick"
    }
    assert extractor._get_border_info(cell) == expected

def test_get_border_info_empty_border_styles(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    cell.border.left.style = None
    cell.border.right.style = None
    cell.border.top.style = None
    cell.border.bottom.style = None

    assert extractor._get_border_info(cell) == {}
