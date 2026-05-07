import pytest
import asyncio
import json
from pathlib import Path
from unittest.mock import MagicMock, AsyncMock, patch
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.fixture
def output_dir(tmp_path):
    d = tmp_path / "output"
    d.mkdir()
    (d / "media").mkdir()
    return d

@pytest.fixture
def extractor(output_dir):
    return KamiExcelExtractor(output_dir=output_dir)

@pytest.mark.asyncio
@patch("litellm.acompletion", new_callable=AsyncMock)
@patch("kami_excel_extractor.core.MetadataExtractor")
@patch("kami_excel_extractor.core.ExcelConverter")
@patch("kami_excel_extractor.core.aiofiles.open")
async def test_pagination_and_caching_flow(mock_aio_open, mock_converter_cls, mock_extractor_cls, mock_litellm, output_dir):
    # Setup mocks
    mock_extractor = mock_extractor_cls.return_value
    mock_extractor.extract.return_value = {
        "sheets": {
            "Sheet1": {"html": "<table></table>", "media": []},
            "Sheet2": {"html": "<table></table>", "media": []}
        }
    }
    
    mock_converter = mock_converter_cls.return_value
    # Sheet1 has 2 pages, Sheet2 has 1 page
    mock_converter.convert.side_effect = [
        [output_dir / "Sheet1-1.png", output_dir / "Sheet1-2.png"], # Sheet1
        output_dir / "Sheet2-1.png"                               # Sheet2
    ]
    
    # Mock image reading
    mock_f = AsyncMock()
    mock_f.read.return_value = b"fake-image"
    mock_ctx = MagicMock()
    mock_ctx.__aenter__ = AsyncMock(return_value=mock_f)
    mock_ctx.__aexit__ = AsyncMock()
    mock_aio_open.return_value = mock_ctx
    
    # Create dummy files so exists() returns True
    (output_dir / "Sheet1-1.png").write_text("dummy")
    (output_dir / "Sheet1-2.png").write_text("dummy")
    (output_dir / "Sheet2-1.png").write_text("dummy")
    
    # Mock LLM response
    mock_choice = MagicMock()
    mock_choice.message.content = '```json\n{"data": []}\n```'
    mock_litellm.return_value = MagicMock(choices=[mock_choice])
    
    # Initialize extractor AFTER patching classes
    extractor = KamiExcelExtractor(output_dir=output_dir)
    
    excel_path = output_dir / "test.xlsx"
    excel_path.write_text("fake excel")
    
    opts = ExtractionOptions(use_visual_context=True, use_cache=True)
    
    # First run: should extract and convert
    result = await extractor.aextract_structured_data(excel_path, options=opts)
    
    assert "Sheet1" in result["sheets"]
    assert "Sheet2" in result["sheets"]
    # Check that convert was called for each sheet
    assert mock_converter.convert.call_count == 2
    
    # Second run: should use cache for raw extraction
    mock_extractor.extract.reset_mock()
    mock_converter.convert.reset_mock()
    
    result2 = await extractor.aextract_structured_data(excel_path, options=opts)
    
    # MetadataExtractor.extract should NOT be called again
    mock_extractor.extract.assert_not_called()
    # But converter SHOULD be called again for images (they aren't fully cached yet)
    assert mock_converter.convert.call_count == 2
    
    assert result2 == result
