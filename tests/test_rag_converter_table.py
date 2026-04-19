import pytest
from kami_excel_extractor.rag_converter import JsonToMarkdownConverter

def test_convert_to_table_basic():
    """Test basic table conversion with standard strings."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Name": "Alice", "Age": "30"},
        {"Name": "Bob", "Age": "25"}
    ]
    keys = ["Name", "Age"]
    result = converter._convert_to_table(data, keys)

    expected = (
        "| Name | Age |\n"
        "| --- | --- |\n"
        "| Alice | 30 |\n"
        "| Bob | 25 |"
    )
    assert result == expected

def test_convert_to_table_missing_keys():
    """Test table conversion when some dictionaries are missing keys."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Name": "Alice", "Age": "30"},
        {"Name": "Bob"} # Missing Age
    ]
    keys = ["Name", "Age"]
    result = converter._convert_to_table(data, keys)

    expected = (
        "| Name | Age |\n"
        "| --- | --- |\n"
        "| Alice | 30 |\n"
        "| Bob |  |"
    )
    assert result == expected

def test_convert_to_table_non_string_values():
    """Test table conversion with non-string values (int, float, None)."""
    converter = JsonToMarkdownConverter()
    data = [
        {"Item": "Apple", "Price": 1.5, "Stock": 10},
        {"Item": "Banana", "Price": 0.5, "Stock": None}
    ]
    keys = ["Item", "Price", "Stock"]
    result = converter._convert_to_table(data, keys)

    # item.get(k, "") returns None if the value is explicitly None,
    # then str(None) becomes "None".
    expected = (
        "| Item | Price | Stock |\n"
        "| --- | --- | --- |\n"
        "| Apple | 1.5 | 10 |\n"
        "| Banana | 0.5 | None |"
    )
    assert result == expected

def test_convert_to_table_non_string_keys():
    """Test table conversion when keys themselves are not strings."""
    converter = JsonToMarkdownConverter()
    data = [
        {1: "One", 2: "Two"},
        {1: "I", 2: "II"}
    ]
    keys = [1, 2]
    result = converter._convert_to_table(data, keys)

    expected = (
        "| 1 | 2 |\n"
        "| --- | --- |\n"
        "| One | Two |\n"
        "| I | II |"
    )
    assert result == expected

def test_convert_to_table_extra_keys_ignored():
    """Test that keys not in the provided keys list are ignored."""
    converter = JsonToMarkdownConverter()
    data = [
        {"ID": 1, "Name": "Alice", "Secret": "hidden"},
        {"ID": 2, "Name": "Bob", "Secret": "hidden"}
    ]
    keys = ["ID", "Name"]
    result = converter._convert_to_table(data, keys)

    expected = (
        "| ID | Name |\n"
        "| --- | --- |\n"
        "| 1 | Alice |\n"
        "| 2 | Bob |"
    )
    assert result == expected

def test_convert_to_table_empty_data():
    """Test with empty data list but provided keys."""
    converter = JsonToMarkdownConverter()
    data = []
    keys = ["A", "B"]
    result = converter._convert_to_table(data, keys)

    expected = (
        "| A | B |\n"
        "| --- | --- |"
    )
    assert result == expected
