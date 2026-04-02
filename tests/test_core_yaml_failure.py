import pytest
from unittest.mock import patch
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

def test_parse_yaml_response_invalid_yaml(extractor):
    """Test exception handling when YAML parsing fails (line 128-130)"""
    invalid_yaml = "key: [unclosed list"
    sheet_name = "Sheet1"

    # Directly call the internal method for focused testing
    result = extractor._parse_yaml_response(invalid_yaml, sheet_name)

    assert "error" in result
    assert result["_raw_yaml"] == invalid_yaml
    # Verify that it caught a YAMLError or similar
    assert "scanner" in result["error"].lower() or "expected" in result["error"].lower()

def test_parse_yaml_response_non_dict(extractor):
    """Test handling of non-dictionary YAML output (line 119)"""
    yaml_str = "- item1\n- item2"
    sheet_name = "Sheet1"

    result = extractor._parse_yaml_response(yaml_str, sheet_name)

    assert result["data"] == ["item1", "item2"]
    assert result["_raw_yaml"] == yaml_str

def test_parse_yaml_response_with_sheets_key(extractor):
    """Test extraction when response contains a 'sheets' key (line 122)"""
    yaml_str = """
sheets:
  Sheet1:
    key1: value1
    key2: value2
  Sheet2:
    other: data
"""
    sheet_name = "Sheet1"

    result = extractor._parse_yaml_response(yaml_str, sheet_name)

    assert result["key1"] == "value1"
    assert result["key2"] == "value2"
    assert result["_raw_yaml"] == yaml_str

def test_parse_yaml_response_with_markdown_blocks(extractor):
    """Test extraction when response is wrapped in markdown code blocks"""
    content = "Here is the data:\n```yaml\nkey: value\n```\nHope it helps!"
    sheet_name = "Sheet1"

    result = extractor._parse_yaml_response(content, sheet_name)

    assert result["key"] == "value"
    assert result["_raw_yaml"] == "key: value"

def test_parse_yaml_response_empty_response(extractor):
    """Test handling of empty YAML response"""
    yaml_str = ""
    sheet_name = "Sheet1"

    result = extractor._parse_yaml_response(yaml_str, sheet_name)

    assert result["_raw_yaml"] == ""
    # safe_load("") returns None, which becomes {} via 'or {}'
    assert len(result) == 1

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_api_failure(mock_litellm, extractor):
    """Test exception handling when litellm.acompletion fails (line 158-160)"""
    mock_litellm.side_effect = Exception("API connection error")

    sheet_name = "Sheet1"
    sheet_content = {"html": "<table></table>"}
    model = "gpt-3.5-turbo"
    system_prompt = "test prompt"
    image_url = "http://example.com/image.png"

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, model, system_prompt, image_url, semaphore=None
    )

    assert name == sheet_name
    assert "error" in result
    assert "API connection error" in result["error"]
    assert result["_raw_yaml"] == ""

@pytest.mark.asyncio
async def test_aprocess_media_summary_missing_file(extractor):
    """Test handling of missing media file in _aprocess_media_summary (line 166)"""
    media_item = {"filename": "non_existent.png"}

    result = await extractor._aprocess_media_summary(media_item, "model", None)

    assert result is None

@pytest.mark.asyncio
async def test_aextract_single_sheet_is_simple_list(extractor):
    """Test is_simple path where structured_data is a list (line 134-141)"""
    sheet_name = "SimpleSheet"
    sheet_content = {
        "is_simple": True,
        "structured_data": [{"col1": "val1"}]
    }

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, "model", "prompt", "url", None
    )

    assert name == sheet_name
    assert result["data"] == [{"col1": "val1"}]
    assert result["_raw_yaml"] == ""

@pytest.mark.asyncio
async def test_aextract_single_sheet_is_simple_dict(extractor):
    """Test is_simple path where structured_data is already a dict (line 134-141)"""
    sheet_name = "SimpleSheetDict"
    sheet_content = {
        "is_simple": True,
        "structured_data": {"custom": "data"}
    }

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, "model", "prompt", "url", None
    )

    assert name == sheet_name
    assert result["custom"] == "data"
    assert result["_raw_yaml"] == ""
