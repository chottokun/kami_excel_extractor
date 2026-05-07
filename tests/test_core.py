import pytest
import json
import asyncio
from pathlib import Path
from datetime import date, datetime
from unittest.mock import MagicMock, patch, AsyncMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("kami_excel_extractor.core.aiofiles.open")
@patch("litellm.acompletion", new_callable=AsyncMock)
async def test_extract_structured_data_basic(mock_litellm, mock_aio_open, mock_extract, mock_convert, extractor, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_extract.return_value = {"sheets": {"Sheet1": {"cells": []}}}
    
    # aiofiles.open のモック (関数自体は同期、戻り値が非同期コンテキストマネージャ)
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    # aiofiles.open(...) returns a context manager
    mock_ctx = MagicMock()
    mock_ctx.__aenter__ = AsyncMock(return_value=mock_f)
    mock_ctx.__aexit__ = AsyncMock()
    mock_aio_open.return_value = mock_ctx
    
    mock_choice = MagicMock()
    mock_choice.message.content = '```json\n{"key": "value"}\n```'
    mock_response = MagicMock()
    mock_response.choices = [mock_choice]
    mock_litellm.return_value = mock_response
    
    # 非同期版を直接呼び出す
    result = await extractor.aextract_structured_data(sample_excel_path)
    
    assert result["sheets"]["Sheet1"]["data"]["key"] == "value"
    assert result["sheets"]["Sheet1"]["_raw_data"] == '{"key": "value"}'

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("litellm.acompletion", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.aiofiles.open")
async def test_extract_structured_data_with_visual_summaries(mock_aio_open, mock_litellm, mock_extract, mock_convert, extractor, sample_excel_path, output_dir):
    mock_convert.return_value = Path("dummy.png")
    
    # aiofiles.open のモック
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    mock_ctx = MagicMock()
    mock_ctx.__aenter__ = AsyncMock(return_value=mock_f)
    mock_ctx.__aexit__ = AsyncMock()
    mock_aio_open.return_value = mock_ctx
    
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
    
    mock_choice_sheet = MagicMock()
    mock_choice_sheet.message.content = '```json\n{"data": {"structured": "data"}}\n```'
    mock_res_sheet = MagicMock()
    mock_res_sheet.choices = [mock_choice_sheet]
    
    mock_choice_media = MagicMock()
    mock_choice_media.message.content = "[画像概要] 要約テキスト"
    mock_res_media = MagicMock()
    mock_res_media.choices = [mock_choice_media]

    # パイプライン順序: 1.要約 & 図表解析(並列) 2.シート解析
    # 画像1枚につき aget_visual_summary と _aprocess_chart_data が呼ばれる
    mock_litellm.side_effect = [mock_res_media, mock_res_media, mock_res_sheet]

    options = ExtractionOptions(include_visual_summaries=True)
    result = await extractor.aextract_structured_data(sample_excel_path, options=options)

    assert "Sheet1" in result["sheets"]
    # 最終結果は _attach_media_to_sheets によって統合される
    assert "media" in result["sheets"]["Sheet1"]
    assert result["sheets"]["Sheet1"]["media"][0]["visual_summary"] == "[画像概要] 要約テキスト"
@pytest.mark.asyncio
@patch("litellm.acompletion", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.aiofiles.open")
async def test_get_visual_summary(mock_aio_open, mock_litellm, extractor, tmp_path):
    img_path = tmp_path / "test.png"
    img_path.write_text("binary data")
    
    # aiofiles.open のモック
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    mock_ctx = MagicMock()
    mock_ctx.__aenter__ = AsyncMock(return_value=mock_f)
    mock_ctx.__aexit__ = AsyncMock()
    mock_aio_open.return_value = mock_ctx
    
    mock_choice = MagicMock()
    mock_choice.message.content = "[画像概要] SUMMARY"
    mock_response = MagicMock()
    mock_response.choices = [mock_choice]
    mock_litellm.return_value = mock_response
    
    summary = await extractor.aget_visual_summary(img_path)
    assert summary == "[画像概要] SUMMARY"

@patch("kami_excel_extractor.core.KamiExcelExtractor.aextract_structured_data", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
def test_extract_rag_chunks(mock_extract_raw, mock_extract_struct, extractor, sample_excel_path):
    mock_extract_raw.return_value = {"sheets": {"Sheet1": {"is_simple": False}}}
    mock_extract_struct.return_value = {"sheets": {"Sheet1": {"_raw_data": '{"data": 1}'}}}
    
    sheet_results, raw_data = extractor.extract_rag_chunks(sample_excel_path)
    
    assert "Sheet1" in sheet_results
    assert "chunks" in sheet_results["Sheet1"]

@patch("litellm.acompletion", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("kami_excel_extractor.core.aiofiles.open")
@pytest.mark.asyncio
async def test_extraction_yaml_parsing_failure(mock_aio_open, mock_extract, mock_convert, mock_litellm, extractor, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_extract.return_value = {"sheets": {"Sheet1": {"html": "<table></table>"}}}
    
    # aiofiles.open のモック
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    mock_aio_open.return_value.__aenter__.return_value = mock_f
    
    mock_choice = MagicMock()
    mock_choice.message.content = "```json\n{\"key\": [unclosed list\n```"
    mock_response = MagicMock()
    mock_response.choices = [mock_choice]
    mock_litellm.return_value = mock_response
    
    result = await extractor.aextract_structured_data(sample_excel_path)
    
    sheet_data = result["sheets"]["Sheet1"]
    assert "error" in sheet_data
    # リトライ後も失敗した場合は _raw_data は空になる実装に合わせる
    assert sheet_data["_raw_data"] == ""

@pytest.mark.asyncio
@patch("litellm.acompletion", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.aiofiles.open")
async def test_aget_visual_summary_failure(mock_open, mock_litellm, extractor, tmp_path):
    img_path = tmp_path / "test_fail.png"
    img_path.write_text("binary data")
    
    # aiofiles.open のモック
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    mock_open.return_value.__aenter__.return_value = mock_f

    # Mock litellm to raise an exception
    mock_litellm.side_effect = Exception("LiteLLM API error")

    summary = await extractor.aget_visual_summary(img_path)
    assert summary == "[画像概要] 解析失敗。"

def test_make_json_serializable(extractor):
    """_make_json_serializable が date/datetime を正しく ISO 形式に変換することを確認"""
    d = date(2023, 1, 1)
    dt = datetime(2023, 1, 1, 12, 0, 0)

    data = {
        "date": d,
        "datetime": dt,
        "nested_dict": {
            "date": d
        },
        "nested_list": [dt, {"date": d}],
        "plain_string": "value",
        "plain_int": 123
    }

    expected = {
        "date": "2023-01-01",
        "datetime": "2023-01-01T12:00:00",
        "nested_dict": {
            "date": "2023-01-01"
        },
        "nested_list": ["2023-01-01T12:00:00", {"date": "2023-01-01"}],
        "plain_string": "value",
        "plain_int": 123
    }

    result = extractor._make_json_serializable(data)
    assert result == expected
