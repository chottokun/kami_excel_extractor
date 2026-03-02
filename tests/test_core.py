import pytest
from pathlib import Path
from src.extractor import extract_comprehensive_map
from src.converter import convert_to_png
import os

@pytest.fixture
def sample_excel():
    # テスト用のサンプルExcelが存在することを確認
    path = Path("sample_hoganshi.xlsx")
    if not path.exists():
        from src.create_sample import create_sample_excel
        create_sample_excel("sample_hoganshi.xlsx")
    return path

def test_metadata_extraction(sample_excel):
    """メタデータ抽出が正常に行われ、期待する座標が含まれているかテスト"""
    output_dir = Path("data/test_output")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    res = extract_comprehensive_map(sample_excel, output_dir)
    assert "sheets" in res
    assert len(res["sheets"]) > 0
    # 「報告者」という文字列がどこかのセルに含まれているか確認
    found = False
    for sheet in res["sheets"].values():
        for cell in sheet["cells"]:
            if "報告者" in cell["v"]:
                found = True
                break
    assert found

def test_image_conversion(sample_excel):
    """ExcelからPNGへの変換が正常に行われるかテスト"""
    output_dir = Path("data/test_output")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # PDF経由の変換ロジックをテスト
    png_path = convert_to_png(sample_excel, output_dir)
    assert png_path.exists()
    assert png_path.suffix == ".png"
    assert png_path.stat().st_size > 0
