import pytest
from kami_excel_extractor.rag_converter import JsonToMarkdownConverter

def test_convert_sheets_data_single():
    """Test _convert_sheets_data with a single sheet."""
    converter = JsonToMarkdownConverter()
    data = {
        "sheets": {
            "Sheet1": {"Key": "Value"}
        }
    }
    # _convert_sheets_data is a private method but we test it directly for coverage
    result = converter._convert_sheets_data(data)
    # Expected:
    # # Sheet1
    #
    # - **Key**: Value
    assert "# Sheet1" in result
    assert "- **Key**: Value" in result
    assert result.startswith("# Sheet1")

def test_convert_sheets_data_multiple():
    """Test _convert_sheets_data with multiple sheets."""
    converter = JsonToMarkdownConverter()
    data = {
        "sheets": {
            "Sheet1": "Content1",
            "Sheet2": "Content2"
        }
    }
    result = converter._convert_sheets_data(data)
    expected_parts = [
        "# Sheet1",
        "Content1",
        "# Sheet2",
        "Content2"
    ]
    for part in expected_parts:
        assert part in result

    # Check separation
    assert "# Sheet1\n\nContent1\n\n# Sheet2\n\nContent2" == result

def test_convert_sheets_data_with_media():
    """Test _convert_sheets_data including media information."""
    converter = JsonToMarkdownConverter()
    data = {
        "sheets": {
            "Sheet1": "Content1"
        },
        "media": [
            {"filename": "image.png", "visual_summary": "Summary"}
        ]
    }
    result = converter._convert_sheets_data(data)
    assert "# Sheet1" in result
    assert "## 関連メディア" in result
    assert "![画像](media/image.png)" in result
    assert "**[画像概要]**: Summary" in result

def test_convert_sheets_data_empty():
    """Test _convert_sheets_data with empty sheets or missing key."""
    converter = JsonToMarkdownConverter()

    # Empty sheets
    assert converter._convert_sheets_data({"sheets": {}}) == ""

    # Missing sheets key
    assert converter._convert_sheets_data({}) == ""

    # Only media
    data = {"media": [{"filename": "image.png"}]}
    result = converter._convert_sheets_data(data)
    assert "## 関連メディア" in result
    # Should not start with H1 header
    assert not result.startswith("# ")

def test_convert_delegation_to_sheets():
    """Test that convert() correctly delegates to _convert_sheets_data at level 1."""
    converter = JsonToMarkdownConverter()
    data = {
        "sheets": {
            "Sheet1": "Content1"
        }
    }
    # Using public convert method
    result = converter.convert(data, level=1)
    assert result == "# Sheet1\n\nContent1"

def test_convert_sheets_data_nested_content():
    """Test _convert_sheets_data where sheet content is complex."""
    converter = JsonToMarkdownConverter()
    data = {
        "sheets": {
            "Sheet1": {
                "Table": [
                    {"A": 1, "B": 2},
                    {"A": 3, "B": 4}
                ]
            }
        }
    }
    result = converter._convert_sheets_data(data)
    assert "# Sheet1" in result
    # When Sheet1 content is a dict, level becomes 2.
    # _convert_dict at level 2 will use level 2 headers if it's a list/dict.
    # In _convert_dict: effective_level = level + 1 if level == 1 else level
    # Since level is 2, effective_level is 2.
    # Header: ## Table
    assert "## Table" in result
    assert "| A | B |" in result
    assert "| 1 | 2 |" in result
