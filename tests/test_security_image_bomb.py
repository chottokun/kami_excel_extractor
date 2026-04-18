import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
import io
from PIL import Image, UnidentifiedImageError
from kami_excel_extractor.extractor import MetadataExtractor, MAX_IMAGE_BYTES

def test_extract_media_skips_large_bytes(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    # Mock an image with size exceeding MAX_IMAGE_BYTES
    mock_img = MagicMock()
    mock_img.anchor._from.row = 0
    mock_img.anchor._from.col = 0

    # 20MB + 1 byte
    large_data = b"a" * (MAX_IMAGE_BYTES + 1)
    mock_img.ref.read.return_value = large_data
    mock_ws._images = [mock_img]

    with patch("PIL.Image.open") as mock_open:
        media_info = extractor._extract_media(mock_ws, "Sheet1")

        # Should be skipped due to byte size check
        assert len(media_info) == 0
        mock_open.assert_not_called()

def test_extract_media_skips_decompression_bomb(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    mock_img = MagicMock()
    mock_img.anchor._from.row = 0
    mock_img.anchor._from.col = 0
    mock_img.ref.read.return_value = b"fake_small_data_but_huge_pixels"
    mock_ws._images = [mock_img]

    with patch("PIL.Image.open") as mock_open:
        # Simulate DecompressionBombError from Pillow
        # In reality, this is raised if pixels > Image.MAX_IMAGE_PIXELS
        mock_open.side_effect = Image.DecompressionBombError("Image size exceeds limit")

        media_info = extractor._extract_media(mock_ws, "Sheet1")

        # Should be kept but with error
        assert len(media_info) == 1
        assert media_info[0]["error"] == "unidentified_format"

        mock_open.assert_called_once()

def test_extract_media_respects_module_level_limit():
    from PIL import Image
    from kami_excel_extractor.extractor import MAX_IMAGE_PIXELS
    assert Image.MAX_IMAGE_PIXELS == MAX_IMAGE_PIXELS
