import pytest
import logging
from unittest.mock import patch, MagicMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("litellm.acompletion")
async def test_aextract_structured_data_conversion_failure(
    mock_litellm, mock_extract, mock_convert, extractor, sample_excel_path, caplog
):
    # Setup mocks
    # 1. Converter fails
    mock_convert.side_effect = Exception("Conversion process failed")

    # 2. Extractor returns dummy sheet data
    mock_extract.return_value = {"sheets": {"Sheet1": {"html": "<table></table>"}}}

    # 3. LLM returns successful response
    mock_choice = MagicMock()
    mock_choice.message.content = '```json\n{"data": [{"item": "success"}]}\n```'
    mock_response = MagicMock()
    mock_response.choices = [mock_choice]
    mock_litellm.return_value = mock_response

    # Execute with visual context enabled (default)
    options = ExtractionOptions(use_visual_context=True)

    with caplog.at_level(logging.WARNING):
        result = await extractor.aextract_structured_data(sample_excel_path, options=options)

    # Verification
    # Check that warning was logged
    assert "Excel-to-image failed: Conversion process failed" in caplog.text
    # Check that extraction still completed
    assert "Sheet1" in result["sheets"]
    assert result["sheets"]["Sheet1"]["data"] == [{"item": "success"}]

    # Verify that mock_convert was indeed called
    mock_convert.assert_called_once()
