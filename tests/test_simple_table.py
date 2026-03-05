import pytest
from pathlib import Path
import openpyxl
from kami_excel_extractor.extractor import MetadataExtractor

@pytest.fixture
def simple_excel(tmp_path):
    path = tmp_path / "simple.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Name", "Value"])
    ws.append([1, "Alpha", "A"])
    ws.append([2, "Beta", "B"])
    wb.save(path)
    return path

@pytest.fixture
def complex_excel(tmp_path):
    path = tmp_path / "complex.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Report"])
    ws.merge_cells("A1:C1")
    ws.append(["ID", "Data"])
    ws.append([100, "Something"])
    wb.save(path)
    return path

def test_is_simple_table(simple_excel, tmp_path):
    extractor = MetadataExtractor(tmp_path)
    wb = openpyxl.load_workbook(simple_excel)
    assert extractor.is_simple_table(wb.active) is True

def test_is_not_simple_table_due_to_merging(complex_excel, tmp_path):
    extractor = MetadataExtractor(tmp_path)
    wb = openpyxl.load_workbook(complex_excel)
    assert extractor.is_simple_table(wb.active) is False

def test_extract_simple_table(simple_excel, tmp_path):
    extractor = MetadataExtractor(tmp_path)
    wb = openpyxl.load_workbook(simple_excel)
    data = extractor.extract_simple_table(wb.active)
    assert len(data) == 2
    assert data[0] == {"ID": "1", "Name": "Alpha", "Value": "A"}
    assert data[1] == {"ID": "2", "Name": "Beta", "Value": "B"}

def test_extract_integration(simple_excel, tmp_path):
    extractor = MetadataExtractor(tmp_path)
    res = extractor.extract(simple_excel)
    sheet_data = res["sheets"]["Sheet"]
    assert sheet_data["is_simple"] is True
    assert "structured_data" in sheet_data
    assert len(sheet_data["structured_data"]) == 2
