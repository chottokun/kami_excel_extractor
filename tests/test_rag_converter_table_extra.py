import pytest
from kami_excel_extractor.rag_converter import JsonToMarkdownConverter

def test_convert_to_table_with_pipes():
    """Test table conversion when values contain pipe characters."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Column": "Value with | pipe"},
        {"Column": "Another | pipe | here"}
    ]
    keys = ["Column"]
    result = converter._convert_to_table(data, keys)

    # Current implementation DOES NOT escape pipes.
    # This test will fail if we expect escaping, or pass if we expect current behavior.
    # To improve the code, we should expect escaped pipes.
    expected = (
        "| Column |\n"
        "| --- |\n"
        "| Value with \\| pipe |\n"
        "| Another \\| pipe \\| here |"
    )
    assert result == expected

def test_convert_to_table_with_newlines():
    """Test table conversion when values contain newline characters."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Column": "Line 1\nLine 2"},
        {"Column": "Single line"}
    ]
    keys = ["Column"]
    result = converter._convert_to_table(data, keys)

    # Standard Markdown tables don't support literal newlines in cells.
    # They should be replaced with <br> or similar.
    expected = (
        "| Column |\n"
        "| --- |\n"
        "| Line 1<br>Line 2 |\n"
        "| Single line |"
    )
    assert result == expected

def test_convert_to_table_with_special_keys():
    """Test table conversion when keys contain special characters."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Key|With|Pipes": "Value"}
    ]
    keys = ["Key|With|Pipes"]
    result = converter._convert_to_table(data, keys)

    # Keys should also have pipes escaped in the header and separator lines.
    expected = (
        "| Key\\|With\\|Pipes |\n"
        "| --- |\n"
        "| Value |"
    )
    assert result == expected

def test_convert_to_table_empty_string_vs_none():
    """Test distinction between empty string and None in table cells."""
    converter = JsonToMarkdownConverter()
    data = [
        {"A": "", "B": None}
    ]
    keys = ["A", "B"]
    result = converter._convert_to_table(data, keys)

    # According to current code:
    # item.get("A", "") -> "" -> str("") -> ""
    # item.get("B", "") -> None -> str(None) -> "None"
    expected = (
        "| A | B |\n"
        "| --- | --- |\n"
        "|  | None |"
    )
    assert result == expected
