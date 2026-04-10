import pytest
from unittest.mock import patch
from kami_excel_extractor.core import KamiExcelExtractor, ExtractionConfig

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

def test_parse_llm_response_invalid_yaml(extractor):
    """Test exception handling when YAML parsing fails"""
    invalid_yaml = "key: [unclosed list"
    sheet_name = "Sheet1"

    # Directly call the internal method for focused testing
    result = extractor._parse_llm_response(invalid_yaml, sheet_name)

    assert "error" in result
    assert result["_raw_data"] == invalid_yaml
    # Verify that it caught a YAMLError or similar
    assert "scanner" in result["error"].lower() or "expected" in result["error"].lower()

def test_parse_llm_response_non_dict(extractor):
    """Test handling of non-dictionary YAML output"""
    yaml_str = "- item1\n- item2"
    sheet_name = "Sheet1"

    result = extractor._parse_llm_response(yaml_str, sheet_name)

    assert result["data"] == ["item1", "item2"]
    assert result["_raw_data"] == yaml_str

def test_parse_llm_response_with_sheets_key(extractor):
    """Test extraction when response contains a 'sheets' key"""
    yaml_str = """
sheets:
  Sheet1:
    data:
      key1: value1
"""
    sheet_name = "Sheet1"

    result = extractor._parse_llm_response(yaml_str, sheet_name)

    assert result["data"]["key1"] == "value1"
    assert result["_raw_data"] == yaml_str

def test_parse_llm_response_with_markdown_blocks(extractor):
    """Test extraction when response is wrapped in markdown code blocks"""
    content = "Here is the data:\n```json\n{\"data\": {\"key\": \"value\"}}\n```\nHope it helps!"
    sheet_name = "Sheet1"

    result = extractor._parse_llm_response(content, sheet_name)

    # Note: SheetData with extra='allow' will keep 'data' field
    assert result["data"]["key"] == "value"
    assert result["_raw_data"] == '{"data": {"key": "value"}}'

def test_parse_llm_response_empty_response(extractor):
    """Test handling of empty YAML response"""
    yaml_str = ""
    sheet_name = "Sheet1"

    result = extractor._parse_llm_response(yaml_str, sheet_name)

    assert result["_raw_data"] == ""
    # Pydantic model with defaults
    assert "data" in result

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_api_failure(mock_litellm, extractor):
    """Test exception handling when litellm.acompletion fails"""
    mock_litellm.side_effect = Exception("API connection error")

    sheet_name = "Sheet1"
    sheet_content = {"html": "<table></table>"}
    config = ExtractionConfig(
        model="gpt-3.5-turbo",
        system_prompt="test prompt",
        image_url="http://example.com/image.png",
        semaphore=None
    )

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, config
    )

    assert name == sheet_name
    assert "error" in result
    assert "API connection error" in result["error"]
    assert result["_raw_data"] == ""

@pytest.mark.asyncio
async def test_aprocess_media_summary_missing_file(extractor):
    """Test handling of missing media file in _aprocess_media_summary"""
    media_item = {"filename": "non_existent.png"}

    result = await extractor._aprocess_media_summary(media_item, "model", None)

    assert result is None

@pytest.mark.asyncio
async def test_aextract_single_sheet_is_simple_list(extractor):
    """Test is_simple path where structured_data is a list"""
    sheet_name = "SimpleSheet"
    sheet_content = {
        "is_simple": True,
        "structured_data": [{"col1": "val1"}]
    }
    config = ExtractionConfig(
        model="model",
        system_prompt="prompt",
        image_url="url",
        semaphore=None
    )

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, config
    )

    assert name == sheet_name
    assert result["data"] == [{"col1": "val1"}]
    assert result["_raw_data"] == ""

@pytest.mark.asyncio
async def test_aextract_single_sheet_is_simple_dict(extractor):
    """Test is_simple path where structured_data is already a dict"""
    sheet_name = "SimpleSheetDict"
    sheet_content = {
        "is_simple": True,
        "structured_data": {"custom": "data"}
    }
    config = ExtractionConfig(
        model="model",
        system_prompt="prompt",
        image_url="url",
        semaphore=None
    )

    name, result = await extractor._aextract_single_sheet(
        sheet_name, sheet_content, config
    )

    assert name == sheet_name
    assert result["custom"] == "data"
    assert result["_raw_data"] == ""
