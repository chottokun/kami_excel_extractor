import pytest
from unittest.mock import MagicMock
from kami_excel_extractor.extractor import MetadataExtractor

def test_get_unit_info_none(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = None
    assert extractor._get_unit_info(cell) is None

def test_get_unit_info_general(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "General"
    assert extractor._get_unit_info(cell) is None

def test_get_unit_info_jpy_symbol(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "¥#,##0"
    assert extractor._get_unit_info(cell) == "JPY"

def test_get_unit_info_jpy_text_upper(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "JPY #,##0"
    assert extractor._get_unit_info(cell) == "JPY"

def test_get_unit_info_jpy_text_lower(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "jpy #,##0"
    assert extractor._get_unit_info(cell) == "JPY"

def test_get_unit_info_usd(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "$#,##0"
    assert extractor._get_unit_info(cell) == "USD"

def test_get_unit_info_percent(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "0.00%"
    assert extractor._get_unit_info(cell) == "PERCENT"

def test_get_unit_info_date_yyyy_mm_dd(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "yyyy-mm-dd"
    assert extractor._get_unit_info(cell) == "DATE"

def test_get_unit_info_date_mm_dd_yyyy(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "MM/DD/YYYY"
    assert extractor._get_unit_info(cell) == "DATE"

def test_get_unit_info_date_yy_mm_dd(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "YY/MM/DD"
    assert extractor._get_unit_info(cell) == "DATE"

def test_get_unit_info_unknown(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    cell = MagicMock()
    cell.number_format = "#,##0"
    assert extractor._get_unit_info(cell) == "#,##0"
