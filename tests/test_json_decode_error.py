import json
import pytest
from unittest.mock import patch
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

def test_parse_llm_response_json_decode_error_fallback(extractor):
    """
    Test that _parse_llm_response correctly handles JSONDecodeError
    in a JSON block and falls back to YAML.
    """
    # Invalid JSON in the json block, valid YAML in a yaml block
    content = """
```json
{ "invalid": [ json
```
```yaml
key: yaml_value
```
"""
    sheet_name = "Sheet1"

    result = extractor._parse_llm_response(content, sheet_name)

    # Since 'key' is not a standard field, it gets wrapped in 'data'
    assert result["data"]["key"] == "yaml_value"
    assert result["_raw_data"] == "key: yaml_value"

def test_parse_llm_response_only_yaml_block(extractor):
    """YAML block only."""
    content = "```yaml\nfoo: bar\n```"
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["foo"] == "bar"
    assert result["_raw_data"] == "foo: bar"

def test_parse_llm_response_no_markdown_blocks(extractor):
    """No markdown blocks, should parse whole content as YAML."""
    content = "just: yaml\ncontent: here"
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["just"] == "yaml"
    assert result["data"]["content"] == "here"
    assert result["_raw_data"] == content

def test_parse_llm_response_both_malformed(extractor):
    """Both blocks are malformed, should return error."""
    content = """
```json
{ "broken"
```
```yaml
[ - not valid yaml
  - : : :
```
"""
    result = extractor._parse_llm_response(content, "Sheet1")
    assert "error" in result
    assert result["_raw_data"] == content

def test_parse_llm_response_sheets_structure(extractor):
    """Test handling of the 'sheets' structure in the response."""
    content = """
```json
{
  "sheets": {
    "Sheet1": {
      "data": [{"a": 1}]
    },
    "Sheet2": {
      "data": [{"b": 2}]
    }
  }
}
```
"""
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"] == [{"a": 1}]

    result2 = extractor._parse_llm_response(content, "Sheet2")
    assert result2["data"] == [{"b": 2}]

def test_parse_llm_response_already_has_data_key(extractor):
    """If it already has 'data' key, it shouldn't wrap again."""
    content = '```json\n{"data": {"nested": "value"}}\n```'
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"] == {"nested": "value"}

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_json_decode_error_mock(mock_acompletion, extractor):
    """
    Test aextract_single_sheet when json.loads raises JSONDecodeError.
    """

    class MockResponse:
        def __init__(self, content):
            self.choices = [type('Choice', (), {'message': type('Message', (), {'content': content})()})]

    content_with_both = '```json\n{"key": "value"}\n```\n```yaml\nkey: yaml_value\n```'
    mock_acompletion.return_value = MockResponse(content_with_both)

    sheet_name = "Sheet1"
    sheet_content = {"html": "<table></table>"}

    # Mock json.loads to fail on the specific JSON string
    original_json_loads = json.loads
    def side_effect(s, **kwargs):
        if s == '{"key": "value"}':
            raise json.JSONDecodeError("mock error", s, 0)
        return original_json_loads(s, **kwargs)

    with patch("json.loads", side_effect=side_effect):
        name, result = await extractor._aextract_single_sheet(
            sheet_name, sheet_content, "model", "prompt", None, None
        )

    assert name == sheet_name
    # Since JSON failed, it should have tried YAML.
    assert result["data"]["key"] == "yaml_value"
