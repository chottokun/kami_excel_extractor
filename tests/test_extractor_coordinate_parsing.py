import pytest
from unittest.mock import MagicMock, patch
from src.kami_excel_extractor.extractor import MetadataExtractor

def test_extract_media_invalid_anchor_logging(tmp_path):
    extractor = MetadataExtractor(tmp_path)
    ws = MagicMock()

    # Mock image with an invalid string anchor
    mock_image = MagicMock()
    mock_image.anchor = "INVALID"
    ws._images = [mock_image]

    with patch("src.kami_excel_extractor.extractor.logger") as mock_logger:
        # We need to mock a few more things to let _extract_media run far enough
        # or just check if logger.warning was called.
        # Since _extract_media will try to access img.ref, let's mock it too.
        mock_image.ref = b"fake_data"

        # We also need to mock PIL.Image.open to avoid actual image processing
        with patch("PIL.Image.open") as mock_pil_open:
            results = extractor._extract_media(ws, "Sheet1")

            # Check if warning was logged
            mock_logger.warning.assert_any_call(
                "Failed to parse coordinate anchor 'INVALID' on sheet Sheet1: invalid literal for int() with base 10: 'D'"
            )

            # Check if coord fell back to "unknown"
            assert results[0]["coord"] == "unknown"

def test_get_bounding_box_invalid_anchor_logging(tmp_path):
    extractor = MetadataExtractor(tmp_path)
    ws = MagicMock()

    # Mock some cells to avoid empty sheet issues
    ws.iter_rows.return_value = []
    ws.merged_cells.ranges = []

    # Mock image with an invalid string anchor
    mock_image = MagicMock()
    mock_image.anchor = "INVALID"
    ws._images = [mock_image]

    with patch("src.kami_excel_extractor.extractor.logger") as mock_logger:
        min_r, max_r, min_c, max_c = extractor._get_bounding_box(ws)

        # Check if warning was logged
        mock_logger.warning.assert_called_with(
            "Failed to parse coordinate anchor 'INVALID' in bounding box calculation: invalid literal for int() with base 10: 'D'"
        )

        # Bounding box should still be (1, 0, 1, 0) or similar depending on implementation
        # Actually it initializes max_r, max_c to 0.
        assert max_r == 0
        assert max_c == 0
