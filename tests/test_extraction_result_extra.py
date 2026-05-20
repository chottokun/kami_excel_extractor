import pytest
from pydantic import ValidationError

from kami_excel_extractor.schema import ExtractionResult, FullExtraction, SheetData


def test_extraction_result_extra_fields_ignored():
    """ExtractionResult should ignore unknown fields and not raise an error."""
    data = {"key": "value"}
    # Pass an unknown field 'unknown_field'
    result = ExtractionResult(data=data, unknown_field="should be ignored")

    # Assert defined fields are present
    assert result.data == data

    # Assert unknown fields are NOT present as attributes
    assert not hasattr(result, "unknown_field")

    # Assert unknown fields are NOT in model_dump()
    dumped = result.model_dump()
    assert "data" in dumped
    assert "unknown_field" not in dumped


def test_extraction_result_model_validate_extra_fields():
    """ExtractionResult.model_validate should ignore unknown fields in the input dict."""
    input_dict = {"data": [{"id": 1}], "extra_info": "extra"}
    result = ExtractionResult.model_validate(input_dict)

    assert result.data == [{"id": 1}]
    assert not hasattr(result, "extra_info")

    dumped = result.model_dump()
    assert "data" in dumped
    assert "extra_info" not in dumped


def test_extraction_result_private_attribute_not_set_from_init():
    """Private attributes (starting with _) should not be set from input dict."""
    # Since _raw_data starts with _, Pydantic treats it as private and doesn't
    # populate it from the input dictionary during validation/initialization.
    # This also confirms it's treated as "extra" or just ignored as private.
    result = ExtractionResult(data={}, _raw_data="should not be set")
    assert result._raw_data is None


def test_sheet_data_extra_fields_ignored():
    """SheetData should also ignore unknown fields as per its config."""
    sheet = SheetData(metadata={"author": "test"}, unknown_field="ignored")
    assert sheet.metadata == {"author": "test"}
    assert not hasattr(sheet, "unknown_field")
    assert "unknown_field" not in sheet.model_dump()


def test_full_extraction_extra_fields_allowed():
    """FullExtraction should ALLOW unknown fields as per its config (extra='allow')."""
    full = FullExtraction(sheets={}, extra_metadata="this should be kept")
    # Since extra='allow', it should be present as an attribute if pydantic version supports it
    # or at least not raise an error.
    # In Pydantic V2 with extra='allow', extra fields are stored.
    assert hasattr(full, "extra_metadata")
    assert full.extra_metadata == "this should be kept"

    dumped = full.model_dump()
    assert dumped["extra_metadata"] == "this should be kept"


def test_extraction_result_no_validation_error_on_extra():
    """Explicitly verify no ValidationError is raised when extra fields are provided."""
    try:
        ExtractionResult(data={}, something_else=123)
    except ValidationError as e:
        pytest.fail(f"ExtractionResult raised ValidationError on extra field: {e}")
