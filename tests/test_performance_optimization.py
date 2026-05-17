import asyncio
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions


@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))


@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("aiofiles.open")
@patch("kami_excel_extractor.utils.CacheManager.aget_file_hash")
@patch("litellm.acompletion")
async def test_bypass_llm_for_simple_table(
    mock_litellm, mock_hash, mock_aio_open, mock_extract, mock_convert, extractor, sample_excel_path
):
    mock_convert.return_value = Path("dummy.png")
    mock_hash.return_value = "fakehash"

    # Mock aiofiles.open
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake binary"
    mock_aio_open.return_value.__aenter__.return_value = mock_f

    # Return one simple sheet and one complex sheet
    mock_extract.return_value = {
        "sheets": {
            "SimpleSheet": {"is_simple": True, "structured_data": [{"col1": "val1"}]},
            "ComplexSheet": {"is_simple": False, "html": "<table>...</table>"},
        }
    }

    mock_response = MagicMock()
    mock_response.choices[0].message.content = '```json\n{"data": {"complex_key": "complex_value"}}\n```'
    mock_litellm.return_value = mock_response

    # Disable cache to keep it simple
    options = ExtractionOptions(use_cache=False)
    result = await extractor.aextract_structured_data(sample_excel_path, options=options)

    # Assert LLM was called only once (for ComplexSheet)
    assert mock_litellm.call_count == 1

    # Verify contents
    assert result["sheets"]["SimpleSheet"]["data"]["data"] == [{"col1": "val1"}]
    assert result["sheets"]["ComplexSheet"]["data"]["complex_key"] == "complex_value"
