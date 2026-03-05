import pytest
import json
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("builtins.open", new_callable=MagicMock)
def test_extract_structured_data_basic(mock_open, mock_extract, mock_convert, extractor, mock_litellm, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_extract.return_value = {"sheets": {"Sheet1": {"cells": []}}}
    
    # open のモックがバイナリデータを返すように設定
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    # LiteLLMのモックが返すJSON
    expected_result = {"key": "value"}
    mock_litellm.return_value.choices[0].message.content = json.dumps(expected_result)
    
    result = extractor.extract_structured_data(sample_excel_path)
    
    assert result == expected_result
    assert mock_litellm.called

@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("kami_excel_extractor.core.KamiExcelExtractor.get_visual_summary")
@patch("builtins.open", new_callable=MagicMock)
def test_extract_structured_data_with_visual_summaries(mock_open, mock_vsum, mock_extract, mock_convert, extractor, mock_litellm, sample_excel_path, output_dir):
    mock_convert.return_value = Path("dummy.png")
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"
    
    # メディアを含むメタデータ
    media_filename = "test_media.png"
    mock_extract.return_value = {
        "sheets": {
            "Sheet1": {
                "media": [{"filename": media_filename}]
            }
        }
    }
    
    # メディアファイルを実際に作成しておく（存在チェックを通るように）
    media_path = output_dir / "media" / media_filename
    media_path.parent.mkdir(parents=True, exist_ok=True)
    media_path.write_text("dummy image")
    
    mock_vsum.return_value = "[画像概要] 要約テキスト"
    mock_litellm.return_value.choices[0].message.content = json.dumps({"data": "structured"})
    
    result = extractor.extract_structured_data(sample_excel_path, include_visual_summaries=True)
    
    assert "media" in result
    assert result["media"][0]["visual_summary"] == "[画像概要] 要約テキスト"
    assert mock_vsum.called

def test_get_visual_summary(extractor, mock_litellm, tmp_path):
    img_path = tmp_path / "test.png"
    img_path.write_text("base64 doesn't care about real PNG content for extraction")
    
    mock_litellm.return_value.choices[0].message.content = "[画像概要] SUMMARY"
    
    summary = extractor.get_visual_summary(img_path)
    assert summary == "[画像概要] SUMMARY"

@patch("kami_excel_extractor.core.KamiExcelExtractor.extract_structured_data")
@patch("kami_excel_extractor.core.DocumentGenerator.generate_pdf")
def test_extract_rag_chunks(mock_pdf, mock_extract, extractor, sample_excel_path):
    mock_extract.return_value = {"content": "data", "media": []}
    
    chunks, md, data = extractor.extract_rag_chunks(sample_excel_path)
    
    assert isinstance(chunks, list)
    assert len(chunks) > 0
    assert md.startswith("- **content**")
    assert data == mock_extract.return_value

def test_extraction_failure(extractor, mock_litellm, sample_excel_path):
    mock_litellm.side_effect = Exception("API Error")
    
    with pytest.raises(Exception):
        extractor.extract_structured_data(sample_excel_path)
