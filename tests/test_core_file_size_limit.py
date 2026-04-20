
import pytest
import asyncio
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor, ExtractionOptions

@pytest.mark.asyncio
async def test_file_size_limit_exceeded(tmp_path):
    extractor = KamiExcelExtractor(output_dir=tmp_path)
    # 実際のエクセルファイルを作成
    excel_path = tmp_path / "test.xlsx"
    excel_path.write_bytes(b"dummy content") # 実際の中身は問わない (statでチェックするため)
    
    # 制限を 0MB に設定 (確実に超える)
    options = ExtractionOptions(max_file_size_mb=0)
    
    with pytest.raises(ValueError) as excinfo:
        await extractor.aextract_structured_data(excel_path, options=options)
    
    assert "exceeds the limit" in str(excinfo.value)
