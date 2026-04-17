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
    # Using 'data' key to match SheetData model more reliably if needed,
    # but since extra='allow' it should work with any key.
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

    # Should fall back to YAML parsing
    assert result["key"] == "yaml_value"
    assert result["_raw_data"] == "key: yaml_value"

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_json_decode_error_mock(mock_acompletion, extractor):
    """
    Test aextract_single_sheet when json.loads raises JSONDecodeError.
    This fulfills the rationale of using a mock.
    """

    class MockResponse:
        def __init__(self, content):
            self.choices = [type('Choice', (), {'message': type('Message', (), {'content': content})()})]

    # Let's use a content that has both blocks for easier success.
    content_with_both = '```json\n{"key": "value"}\n```\n```yaml\nkey: yaml_value\n```'
    mock_acompletion.return_value = MockResponse(content_with_both)

    sheet_name = "Sheet1"
    sheet_content = {"html": "<table></table>"}

    # We want to mock json.loads but only when it's called with the expected string
    original_json_loads = json.loads
    def side_effect(s, **kwargs):
        if s == '{"key": "value"}':
            raise json.JSONDecodeError("mock error", s, 0)
        return original_json_loads(s, **kwargs)

    with patch("json.loads", side_effect=side_effect):
        name, result = await extractor._aextract_single_sheet(
            sheet_name, sheet_content, "model", "prompt", "url", None
        )

    assert name == sheet_name
    # Since JSON failed, it should have tried YAML.
    assert result["key"] == "yaml_value"
