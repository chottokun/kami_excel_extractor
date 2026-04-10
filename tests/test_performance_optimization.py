import pytest
import asyncio
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.core import KamiExcelExtractor

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("builtins.open", new_callable=MagicMock)
@patch("litellm.acompletion")
async def test_bypass_llm_for_simple_table(mock_litellm, mock_open, mock_extract, mock_convert, extractor, sample_excel_path):
    mock_convert.return_value = Path("dummy.png")
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"

    # Return one simple sheet and one complex sheet
    mock_extract.return_value = {
        "sheets": {
            "SimpleSheet": {
                "is_simple": True,
                "structured_data": [{"col1": "val1"}]
            },
            "ComplexSheet": {
                "is_simple": False,
                "html": "<table>...</table>"
            }
        }
    }

    mock_response = MagicMock()
    mock_response.choices[0].message.content = "```yaml\ncomplex_key: complex_value\n```"
    mock_litellm.return_value = mock_response

    result = await extractor.aextract_structured_data(sample_excel_path)

    # Assert LLM was called only once (for ComplexSheet)
    assert mock_litellm.call_count == 1

    # Verify contents
    assert result["sheets"]["SimpleSheet"]["data"] == [{"col1": "val1"}]
    assert result["sheets"]["SimpleSheet"]["_raw_data"] == ""
    assert result["sheets"]["ComplexSheet"]["complex_key"] == "complex_value"
