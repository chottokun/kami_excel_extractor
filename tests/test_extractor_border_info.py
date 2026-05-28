from unittest.mock import MagicMock

from openpyxl.styles import Border, Side

from kami_excel_extractor.extractor import MetadataExtractor


def test_get_border_info_no_border(tmp_path):
    """境界線が設定されていないセルのテスト"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.border = None

    borders = extractor._get_border_info(cell)
    assert borders == {}


def test_get_border_info_partial_border(tmp_path):
    """一部の辺のみに境界線が設定されているセルのテスト"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    # 本物のBorderとSideを使用する
    left_side = Side(style="thin")
    top_side = Side(style="thick")
    # right_sideとbottom_sideはスタイルなし
    cell.border = Border(left=left_side, top=top_side)

    borders = extractor._get_border_info(cell)
    assert borders == {"left": "thin", "top": "thick"}


def test_get_border_info_full_border(tmp_path):
    """すべての辺に境界線が設定されているセルのテスト"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    cell.border = Border(
        left=Side(style="thin"),
        right=Side(style="medium"),
        top=Side(style="thick"),
        bottom=Side(style="dashed"),
    )

    borders = extractor._get_border_info(cell)
    assert borders == {
        "left": "thin",
        "right": "medium",
        "top": "thick",
        "bottom": "dashed",
    }


def test_get_border_info_border_no_style(tmp_path):
    """Borderオブジェクトは存在するが、スタイルが空のセルのテスト"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()

    cell.border = Border(
        left=Side(style=None),
        right=Side(style=None),
        top=Side(style=None),
        bottom=Side(style=None),
    )

    borders = extractor._get_border_info(cell)
    assert borders == {}
