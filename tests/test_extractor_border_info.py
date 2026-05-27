from unittest.mock import MagicMock
from kami_excel_extractor.extractor import MetadataExtractor

def test_get_border_info_no_border(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.border = None

    borders = extractor._get_border_info(cell)
    assert borders == {}

def test_get_border_info_partial_border(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    # Mocking border sides
    left_side = MagicMock()
    left_side.style = "thin"

    right_side = MagicMock()
    right_side.style = None

    top_side = MagicMock()
    top_side.style = "thick"

    bottom_side = None

    cell.border.left = left_side
    cell.border.right = right_side
    cell.border.top = top_side
    cell.border.bottom = bottom_side

    borders = extractor._get_border_info(cell)
    assert borders == {"left": "thin", "top": "thick"}

def test_get_border_info_full_border(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    styles = {
        "left": "thin",
        "right": "medium",
        "top": "thick",
        "bottom": "dashed"
    }

    for side, style in styles.items():
        side_mock = MagicMock()
        side_mock.style = style
        setattr(cell.border, side, side_mock)

    borders = extractor._get_border_info(cell)
    assert borders == styles

def test_get_border_info_border_no_style(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    for side in ["left", "right", "top", "bottom"]:
        side_mock = MagicMock()
        side_mock.style = None
        setattr(cell.border, side, side_mock)

    borders = extractor._get_border_info(cell)
    assert borders == {}
