import pytest
from unittest.mock import patch
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

def test_parse_llm_response_validation_exception(extractor):
    """
    Test that _parse_llm_response correctly handles a generic Exception
    during Pydantic validation (SheetData instantiation).
    """
    content = '{"data": {"some": "data"}}'
    sheet_name = "Sheet1"

    # Mock SheetData to raise a generic Exception
    with patch("kami_excel_extractor.core.SheetData", side_effect=Exception("Mocked validation error")):
        result = extractor._parse_llm_response(content, sheet_name)

    assert "error" in result
    assert "Validation failed: Mocked validation error" in result["error"]
    assert result["_raw_data"] == '{"data": {"some": "data"}}'

def test_parse_llm_response_validation_error_pydantic(extractor):
    """
    Test that _parse_llm_response correctly handles a Pydantic ValidationError
    (which is a subclass of Exception).
    """
    # Providing data that would normally fail if we weren't using extra='allow'
    # but here we'll force a ValidationError by mocking.
    content = '{"data": "invalid"}'
    sheet_name = "Sheet1"

    # We use ValueError as a representative of validation errors that can occur
    with patch("kami_excel_extractor.core.SheetData", side_effect=ValueError("Invalid data format")):
        result = extractor._parse_llm_response(content, sheet_name)

    assert "error" in result
    assert "Validation failed: Invalid data format" in result["error"]
    assert result["_raw_data"] == '{"data": "invalid"}'
