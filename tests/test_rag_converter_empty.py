import pytest
from kami_excel_extractor.rag_converter import JsonToMarkdownConverter

def test_convert_none():
    converter = JsonToMarkdownConverter()
    assert converter.convert(None) == "None"

def test_convert_empty_dict():
    converter = JsonToMarkdownConverter()
    assert converter.convert({}) == ""

def test_convert_empty_sheets():
    converter = JsonToMarkdownConverter()
    assert converter.convert({"sheets": {}}) == ""

def test_convert_sheet_with_none():
    converter = JsonToMarkdownConverter()
    # Improved behavior: skip sheets with no content
    result = converter.convert({"sheets": {"Sheet1": None}})
    assert result == ""

def test_convert_sheet_with_empty_dict():
    converter = JsonToMarkdownConverter()
    # Improved behavior: skip sheets with no content
    result = converter.convert({"sheets": {"Sheet1": {}}})
    assert result == ""

def test_convert_sheet_with_empty_list():
    converter = JsonToMarkdownConverter()
    # Improved behavior: skip sheets with no content
    result = converter.convert({"sheets": {"Sheet1": []}})
    assert result == ""

def test_convert_nested_empty_dict():
    converter = JsonToMarkdownConverter()
    # Improved behavior: skip keys with no content
    assert converter.convert({"key": {}}) == ""

def test_convert_only_empty_media():
    converter = JsonToMarkdownConverter()
    assert converter.convert({"media": []}) == ""
