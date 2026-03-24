import pytest
import json
import asyncio
from pathlib import Path
from unittest.mock import MagicMock, patch, AsyncMock
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("builtins.open", new_callable=MagicMock)
@patch("litellm.acompletion")
async def test_extract_structured_data_basic(mock_litellm, mock_open, mock_extract, mock_convert, extractor, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_extract.return_value = {"sheets": {"Sheet1": {"cells": []}}}
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    mock_response = MagicMock()
    mock_response.choices[0].message.content = "```yaml\nkey: value\n```"
    mock_litellm.return_value = mock_response
    
    # 非同期版を直接呼び出す
    result = await extractor.aextract_structured_data(sample_excel_path)
    
    assert result["sheets"]["Sheet1"]["key"] == "value"
    assert result["sheets"]["Sheet1"]["_raw_yaml"] == "key: value"

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("litellm.acompletion")
@patch("builtins.open", new_callable=MagicMock)
async def test_extract_structured_data_with_visual_summaries(mock_open, mock_litellm, mock_extract, mock_convert, extractor, sample_excel_path, output_dir):
    mock_convert.return_value = Path("dummy.png")
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    media_filename = "Sheet1_img_A1_0.png"
    media_path = output_dir / "media" / media_filename
    media_path.parent.mkdir(parents=True, exist_ok=True)
    media_path.write_text("dummy image")
    
    mock_extract.return_value = {
        "sheets": {
            "Sheet1": {
                "media": [{"filename": media_filename, "coord": "A1"}]
            }
        }
    }
    
    mock_res_sheet = MagicMock()
    mock_res_sheet.choices[0].message.content = "```yaml\ndata: structured\n```"
    mock_res_media = MagicMock()
    mock_res_media.choices[0].message.content = "[画像概要] 要約テキスト"
    mock_litellm.side_effect = [mock_res_sheet, mock_res_media]
    
    result = await extractor.aextract_structured_data(sample_excel_path, include_visual_summaries=True)
    
    assert "Sheet1" in result["sheets"]
    assert "media" in result["sheets"]["Sheet1"]
    assert result["sheets"]["Sheet1"]["media"][0]["visual_summary"] == "[画像概要] 要約テキスト"

@pytest.mark.asyncio
@patch("litellm.acompletion")
@patch("builtins.open", new_callable=MagicMock)
async def test_get_visual_summary(mock_open, mock_litellm, extractor, tmp_path):
    img_path = tmp_path / "test.png"
    img_path.write_text("binary data")
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    mock_response = MagicMock()
    mock_response.choices[0].message.content = "[画像概要] SUMMARY"
    mock_litellm.return_value = mock_response
    
    summary = await extractor.aget_visual_summary(img_path)
    assert summary == "[画像概要] SUMMARY"

@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.KamiExcelExtractor.aextract_structured_data")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
def test_extract_rag_chunks(mock_extract_raw, mock_aextract_struct, mock_convert, extractor, sample_excel_path):
    # aextract_rag_chunks calls aextract_structured_data (async)
    mock_convert.return_value = Path("dummy.png")
    mock_extract_raw.return_value = {"sheets": {"Sheet1": {"is_simple": False}}}

    async def mock_aextract(*args, **kwargs):
        return {"sheets": {"Sheet1": {"_raw_yaml": "data: 1"}}}

    mock_aextract_struct.side_effect = mock_aextract
    
    sheet_results, raw_data = extractor.extract_rag_chunks(sample_excel_path)
    
    assert "Sheet1" in sheet_results
    assert "chunks" in sheet_results["Sheet1"]

@patch("litellm.acompletion")
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("builtins.open", new_callable=MagicMock)
@pytest.mark.asyncio
async def test_extraction_yaml_parsing_failure(mock_open, mock_extract, mock_convert, mock_litellm, extractor, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_extract.return_value = {"sheets": {"Sheet1": {"html": "<table></table>"}}}
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    mock_response = MagicMock()
    mock_response.choices[0].message.content = "```yaml\nkey: [unclosed list\n```"
    mock_litellm.return_value = mock_response
    
    result = await extractor.aextract_structured_data(sample_excel_path)
    
    sheet_data = result["sheets"]["Sheet1"]
    assert "error" in sheet_data
    assert "key: [unclosed list" in sheet_data["_raw_yaml"]
