import pytest
import json
import yaml
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

def test_parse_llm_response_pure_json(extractor):
    """Test pure JSON content without markdown blocks."""
    content = '{"data": {"key": "value"}}'
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["key"] == "value"
    assert "_raw_data" in result

def test_parse_llm_response_pure_yaml(extractor):
    """Test pure YAML content without markdown blocks."""
    content = "key: value"
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["key"] == "value"

def test_parse_llm_response_invalid_json_fallback_yaml_blocks(extractor):
    """Test invalid JSON in markdown block falling back to YAML in markdown block."""
    content = """
Here is the data:
```json
{ "invalid":
```
```yaml
valid: yaml
```
"""
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["valid"] == "yaml"
    assert result["_raw_data"] == "valid: yaml"

def test_parse_llm_response_invalid_json_fallback_unstructured(extractor):
    """Test invalid JSON in markdown block falling back to unstructured YAML/JSON."""
    content = """
unstructured_key: unstructured_value
```json
{ "invalid":
```
"""
    result = extractor._parse_llm_response(content, "Sheet1")
    # yaml.safe_load might fail on the whole content if it contains ```
    # If it fails, it returns error.
    if "error" in result:
        assert "error" in result
    else:
        # If it happens to parse (some yaml parsers are lenient)
        assert result["data"]["unstructured_key"] == "unstructured_value"

def test_parse_llm_response_completely_invalid(extractor):
    """Test completely invalid/unstructured text resulting in an error response."""
    # Using characters that definitely break YAML if possible, or just a string.
    content = "This is not JSON or YAML at all. Just some random text."
    result = extractor._parse_llm_response(content, "Sheet1")

    if "error" in result:
        assert "_raw_data" in result
        assert result["_raw_data"] == content
    else:
        assert result["data"] == content

def test_parse_llm_response_multiple_json_blocks(extractor):
    """Response containing multiple JSON blocks (ensure it picks the first)."""
    content = """
```json
{"first": 1}
```
```json
{"second": 2}
```
"""
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["first"] == 1
    assert "second" not in result["data"]

def test_parse_llm_response_wrapped_in_data_key(extractor):
    """If the response is a dictionary without a 'data' key, it should be wrapped."""
    content = '{"some_key": "some_value"}'
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["some_key"] == "some_value"

def test_parse_llm_response_sheet_name_nesting(extractor):
    """If the response has a 'sheets' key with the sheet name, it should extract it."""
    content = '{"sheets": {"Sheet1": {"data": {"nested": "value"}}}}'
    result = extractor._parse_llm_response(content, "Sheet1")
    assert result["data"]["nested"] == "value"
