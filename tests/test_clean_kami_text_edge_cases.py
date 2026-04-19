import pytest
from kami_excel_extractor.utils import clean_kami_text

def test_clean_kami_text_non_string():
    """clean_kami_text returns non-string inputs unmodified."""
    assert clean_kami_text(None) is None
    assert clean_kami_text(123) == 123
    assert clean_kami_text(45.6) == 45.6
    assert clean_kami_text(True) is True
    assert clean_kami_text(False) is False
    assert clean_kami_text(["a", "b"]) == ["a", "b"]
    assert clean_kami_text({"key": "value"}) == {"key": "value"}

def test_clean_kami_text_empty_string():
    """clean_kami_text handles empty strings correctly."""
    assert clean_kami_text("") == ""
    assert clean_kami_text("   ") == ""
