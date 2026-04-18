import pytest
from unittest.mock import MagicMock, patch, AsyncMock
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(api_key="fake_key", output_dir=str(output_dir))

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_retry_on_exception(mock_litellm, extractor):
    """Test that _aextract_single_sheet retries when an Exception occurs in litellm.acompletion"""
    # First call raises Exception, second call succeeds
    mock_success_response = MagicMock()
    mock_success_response.choices = [MagicMock(message=MagicMock(content='```json\n{"data": ["success"]}\n```'))]

    mock_litellm.side_effect = [
        Exception("First attempt failed"),
        mock_success_response
    ]

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
    assert mock_litellm.call_count == 2

@pytest.mark.asyncio
@patch("litellm.acompletion")
async def test_aextract_single_sheet_all_retries_fail(mock_litellm, extractor):
    """Test that _aextract_single_sheet returns error after all retries fail due to Exceptions"""
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
    assert "Failed after 1 retries" in result["error"]
    assert "Persistent error" in result["error"]
    assert mock_litellm.call_count == 2

@pytest.mark.asyncio
@patch("kami_excel_extractor.core.ExcelConverter.convert")
@patch("kami_excel_extractor.core.MetadataExtractor.extract")
@patch("litellm.acompletion")
@patch("builtins.open", new_callable=MagicMock)
async def test_aextract_structured_data_partial_failure(mock_open, mock_litellm, mock_extract, mock_convert, extractor, sample_excel_path):
    """Test that aextract_structured_data handles cases where some sheets fail and others succeed"""
    mock_convert.return_value = Path("dummy.png")
    # Need to mock open for the image encoding part
    mock_open.return_value.__enter__.return_value.read.return_value = b"fake binary"

    # Mocking two sheets
    mock_extract.return_value = {
        "sheets": {
            "SheetSucceed": {"html": "<table></table>"},
            "SheetFail": {"html": "<table></table>"}
        }
    }

    # Responses:
    # SheetSucceed: Success
    # SheetFail: Exception, then Exception (all retries fail)

    mock_success_res = MagicMock()
    mock_success_res.choices = [MagicMock(message=MagicMock(content='```json\n{"data": ["ok"]}\n```'))]

    def side_effect(*args, **kwargs):
        # Determine which sheet is being processed based on messages
        messages = kwargs.get('messages', [])
        content = ""
        for msg in messages:
            if msg['role'] == 'user':
                if isinstance(msg['content'], list):
                    content = msg['content'][0]['text']
                else:
                    content = msg['content']
                break

        if "SheetSucceed" in content:
            return mock_success_res
        else:
            raise Exception("SheetFail error")

    mock_litellm.side_effect = side_effect

    result = await extractor.aextract_structured_data(sample_excel_path)

    assert "SheetSucceed" in result["sheets"]
    assert "SheetFail" in result["sheets"]

    assert result["sheets"]["SheetSucceed"]["data"] == ["ok"]
    assert "error" in result["sheets"]["SheetFail"]
    assert "Failed after 1 retries" in result["sheets"]["SheetFail"]["error"]
