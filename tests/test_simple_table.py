import pytest
import openpyxl
from datetime import date, datetime
from kami_excel_extractor.extractor import MetadataExtractor

@pytest.fixture
def extractor(tmp_path):
    return MetadataExtractor(output_dir=tmp_path)

def test_is_simple_table_with_merged_cells(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    ws["A2"] = "D1"
    ws.merge_cells("A1:B1")
    assert extractor.is_simple_table(ws) is False

def test_is_simple_table_insufficient_rows(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    # Only 1 row
    assert extractor.is_simple_table(ws) is False

def test_is_simple_table_insufficient_headers(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = None
    ws["A2"] = "D1"
    # Only 1 non-None header
    assert extractor.is_simple_table(ws) is False

def test_is_simple_table_valid(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    ws["A2"] = "D1"
    ws["B2"] = "D2"
    assert extractor.is_simple_table(ws) is True

def test_extract_simple_table_mapping(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "ID"
    ws["B1"] = "Name"
    ws["A2"] = 1
    ws["B2"] = "Alice"
    ws["A3"] = 2
    ws["B3"] = "Bob"

    data = extractor.extract_simple_table(ws)
    assert data == [
        {"ID": 1, "Name": "Alice"},
        {"ID": 2, "Name": "Bob"}
    ]

def test_extract_simple_table_default_headers(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "ID"
    ws["B1"] = None # Should be Column2
    ws["A2"] = 1
    ws["B2"] = "Val"

    data = extractor.extract_simple_table(ws)
    assert data == [{"ID": 1, "Column2": "Val"}]

def test_extract_simple_table_date_conversion(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Event"
    ws["B1"] = "Date"
    ws["A2"] = "Birthday"
    ws["B2"] = date(1990, 1, 1)
    ws["A3"] = "Now"
    ws["B3"] = datetime(2023, 10, 27, 10, 0, 0)

    data = extractor.extract_simple_table(ws)
    assert data[0]["Date"] == "1990-01-01"
    assert data[1]["Date"] == "2023-10-27T10:00:00"

def test_extract_simple_table_none_to_empty_string(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    ws["A2"] = "Val1"
    ws["B2"] = None

    data = extractor.extract_simple_table(ws)
    assert data == [{"H1": "Val1", "H2": ""}]

def test_extract_simple_table_skips_empty_rows(extractor):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "H1"
    ws["B1"] = "H2"
    ws["A2"] = "Val1"
    ws["B2"] = "Val2"
    # Row 3 is empty
    ws["A4"] = "Val3"
    ws["B4"] = "Val4"

    data = extractor.extract_simple_table(ws)
    assert len(data) == 2
    assert data[0]["H1"] == "Val1"
    assert data[1]["H1"] == "Val3"
