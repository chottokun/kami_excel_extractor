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
@patch("litellm.acompletion")
async def test_aextract_single_sheet_retry_on_exception(mock_litellm, extractor):
    """Test that _aextract_single_sheet retries when an Exception occurs in litellm.acompletion"""
    # Force fast retries for tests
    with patch("asyncio.sleep", return_value=None):
        # First call raises Exception, second call succeeds
        mock_success_response = MagicMock()
        mock_success_response.choices = [MagicMock(message=MagicMock(content='```json\n{"data": ["success"]}\n```'))]

        mock_litellm.side_effect = [Exception("First attempt failed"), mock_success_response]

        sheet_name = "Sheet1"
        sheet_content = {"html": "<table></table>"}
        model = "gpt-3.5-turbo"
        system_prompt = "test prompt"
        image_url = None

        name, result = await extractor._aextract_single_sheet(
            sheet_name, sheet_content, model, system_prompt, image_url, semaphore=None
        )

        assert name == sheet_name
        assert result["data"] == ["success"]
        # _awith_retry might have retried internally, or _aextract_single_sheet's loop retried.
        # With default max_retries=3 in _awith_retry, it would retry 3 times before failing to the outer loop.
        # Here we just want to ensure it eventually succeeds.
        assert mock_litellm.call_count >= 2


@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_all_retries_fail(mock_litellm, extractor):
    """Test that _aextract_single_sheet returns error after all retries fail due to Exceptions"""
    with patch("asyncio.sleep", return_value=None):
        mock_litellm.side_effect = Exception("Persistent error")

        sheet_name = "Sheet1"
        sheet_content = {"html": "<table></table>"}
        model = "gpt-3.5-turbo"
        system_prompt = "test prompt"
        image_url = None

        name, result = await extractor._aextract_single_sheet(
            sheet_name, sheet_content, model, system_prompt, image_url, semaphore=None
        )

        assert name == sheet_name
        assert "error" in result
        assert "Persistent error" in result["error"]


@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("litellm.acompletion")
@patch("aiofiles.open")
@patch("kami_excel_extractor.utils.CacheManager.aget_file_hash")
async def test_aextract_structured_data_partial_failure(
    mock_hash, mock_aio_open, mock_litellm, mock_extract, mock_convert, extractor, sample_excel_path
):
    """Test that aextract_structured_data handles cases where some sheets fail and others succeed"""
    # Use fast retries and disable cache
    extractor.cache.clear()
    mock_hash.return_value = "fakehash"

    with patch("asyncio.sleep", return_value=None):
        mock_convert.return_value = Path("dummy.png")

        # Mock aiofiles.open for image encoding
        mock_f = AsyncMock()
        mock_f.read.return_value = b"fake binary"
        mock_aio_open.return_value.__aenter__.return_value = mock_f

        # Mocking two sheets with unique HTML to avoid cache collisions
        mock_extract.return_value = {
            "sheets": {
                "SheetSucceed": {"html": "<!-- succeed --><table></table>"},
                "SheetFail": {"html": "<!-- fail --><table></table>"},
            }
        }

        mock_success_res = MagicMock()
        mock_success_res.choices = [MagicMock(message=MagicMock(content='```json\n{"data": ["ok"]}\n```'))]

        async def async_side_effect(*args, **kwargs):
            messages = kwargs.get("messages", [])
            content = str(messages)
            if "SheetSucceed" in content:
                return mock_success_res
            else:
                raise Exception("SheetFail error")

        mock_litellm.side_effect = async_side_effect

        # Disable cache via options
        options = ExtractionOptions(use_cache=False)
        result = await extractor.aextract_structured_data(sample_excel_path, options=options)

        assert "SheetSucceed" in result["sheets"]
        assert "SheetFail" in result["sheets"]

        assert result["sheets"]["SheetSucceed"]["data"] == ["ok"]
        assert "error" in result["sheets"]["SheetFail"]
