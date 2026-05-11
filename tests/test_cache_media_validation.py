import pytest
import json
from pathlib import Path
from unittest.mock import MagicMock, patch, AsyncMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.mark.asyncio
async def test_cache_media_missing_triggers_reextraction(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    extractor = KamiExcelExtractor(output_dir=output_dir, api_key="fake")

    excel_path = tmp_path / "test.xlsx"
    excel_path.write_text("dummy")

    # Mock cache to return a hit for raw extraction
    cached_data = {
        "sheets": {
            "Sheet1": {
                "media": [{"filename": "missing.png"}]
            }
        }
    }

    with patch.object(extractor.cache, 'get_file_hash', return_value="hash123"), \
         patch.object(extractor.cache, 'get_raw_extraction', return_value=json.dumps(cached_data)), \
         patch.object(extractor.extractor, 'extract', side_effect=Exception("Re-extraction triggered")) as mock_extract:

        # The file output/media/missing.png does NOT exist.

        options = ExtractionOptions(use_cache=True)

        # We expect it to try to re-extract because media is missing
        with pytest.raises(Exception, match="Re-extraction triggered"):
            await extractor.aextract_structured_data(excel_path, options=options)

        mock_extract.assert_called_once()

@pytest.mark.asyncio
async def test_cache_media_exists_uses_cache(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    media_dir = output_dir / "media"
    media_dir.mkdir()
    (media_dir / "exists.png").write_text("dummy")

    extractor = KamiExcelExtractor(output_dir=output_dir, api_key="fake")

    excel_path = tmp_path / "test.xlsx"
    excel_path.write_text("dummy")

    cached_data = {
        "sheets": {
            "Sheet1": {
                "media": [{"filename": "exists.png"}],
                "html": "<table></table>"
            }
        }
    }

    with patch.object(extractor.cache, 'get_file_hash', return_value="hash123"), \
         patch.object(extractor.cache, 'get_raw_extraction', return_value=json.dumps(cached_data)), \
         patch.object(extractor.extractor, 'extract') as mock_extract:

        options = ExtractionOptions(use_cache=True, use_visual_context=False, include_visual_summaries=False)

        # Mocking _aextract_single_sheet to avoid actual LLM call for this test
        with patch.object(extractor, '_aextract_single_sheet', AsyncMock(return_value=("Sheet1", {"data": {}}))):
            await extractor.aextract_structured_data(excel_path, options=options)

        # extract should NOT be called because cache was valid
        mock_extract.assert_not_called()
