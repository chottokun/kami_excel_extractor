import pytest
import logging
from pathlib import Path
from unittest.mock import MagicMock, patch
from PIL import Image, UnidentifiedImageError
from kami_excel_extractor.extractor import MetadataExtractor

def test_extract_media_os_error(tmp_path, caplog):
    """OSError during Image.open should be caught, logged, and skipped."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    mock_img = MagicMock()
    mock_img.anchor._from.row = 0
    mock_img.anchor._from.col = 0
    mock_img.ref.read.return_value = b"corrupt_data"
    mock_ws._images = [mock_img]

    with patch("PIL.Image.open") as mock_open:
        mock_open.side_effect = OSError("Internal library error")

        with caplog.at_level(logging.WARNING):
            media_info = extractor._extract_media(mock_ws, "Sheet1")

        assert media_info == []
        assert "Failed to extract image at A1 on sheet Sheet1: Internal library error" in caplog.text

def test_extract_media_attribute_error(tmp_path, caplog):
    """AttributeError during Image processing should be caught, logged, and skipped."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    mock_img = MagicMock()
    mock_img.anchor._from.row = 1
    mock_img.anchor._from.col = 1
    mock_img.ref.read.return_value = b"some_data"
    mock_ws._images = [mock_img]

    with patch("PIL.Image.open") as mock_open:
        mock_pillow_img = MagicMock(spec=Image.Image)
        mock_pillow_img.__enter__.return_value = mock_pillow_img
        mock_pillow_img.mode = "RGB"
        mock_pillow_img.save.side_effect = AttributeError("Mock attribute error")
        mock_open.return_value = mock_pillow_img

        with caplog.at_level(logging.WARNING):
            media_info = extractor._extract_media(mock_ws, "Sheet2")

        assert media_info == []
        assert "Failed to extract image at B2 on sheet Sheet2: Mock attribute error" in caplog.text

def test_extract_media_mixed_success_and_failure(tmp_path, caplog):
    """Extraction should continue for other images when one fails with a caught exception."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    # Image 1: Success
    mock_img1 = MagicMock()
    mock_img1.anchor._from.row = 0
    mock_img1.anchor._from.col = 0
    mock_img1.ref.read.return_value = b"data1"

    # Image 2: Failure (UnidentifiedImageError)
    mock_img2 = MagicMock()
    mock_img2.anchor._from.row = 1
    mock_img2.anchor._from.col = 1
    mock_img2.ref.read.return_value = b"data2"

    # Image 3: Success
    mock_img3 = MagicMock()
    mock_img3.anchor._from.row = 2
    mock_img3.anchor._from.col = 2
    mock_img3.ref.read.return_value = b"data3"

    mock_ws._images = [mock_img1, mock_img2, mock_img3]

    with patch("PIL.Image.open") as mock_open:
        mock_pillow_ok1 = MagicMock(spec=Image.Image)
        mock_pillow_ok1.mode = "RGB"
        mock_pillow_ok1.__enter__.return_value = mock_pillow_ok1

        mock_pillow_ok3 = MagicMock(spec=Image.Image)
        mock_pillow_ok3.mode = "RGB"
        mock_pillow_ok3.__enter__.return_value = mock_pillow_ok3

        mock_open.side_effect = [
            mock_pillow_ok1,
            UnidentifiedImageError("Unknown format"),
            mock_pillow_ok3
        ]

        with caplog.at_level(logging.WARNING):
            media_info = extractor._extract_media(mock_ws, "Sheet1")

        assert len(media_info) == 2
        assert media_info[0]["coord"] == "A1"
        assert media_info[1]["coord"] == "C3"
        assert "Failed to extract image at B2 on sheet Sheet1: Unknown format" in caplog.text
