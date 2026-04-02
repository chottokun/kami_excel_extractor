import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
import io
from PIL import Image, UnidentifiedImageError
from kami_excel_extractor.extractor import MetadataExtractor

def test_extract_media_success(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Mock worksheet
    mock_ws = MagicMock()

    # Mock images
    mock_img1 = MagicMock()
    mock_img1.anchor._from.row = 0
    mock_img1.anchor._from.col = 0
    mock_img1.ref.read.return_value = b"fake_image_data_1"

    mock_img2 = MagicMock()
    mock_img2.anchor._from.row = 1
    mock_img2.anchor._from.col = 1
    mock_img2.ref.getvalue.return_value = b"fake_image_data_2"
    # Ensure it doesn't have 'read' to trigger 'getvalue'
    if hasattr(mock_img2.ref, "read"):
        del mock_img2.ref.read

    mock_ws._images = [mock_img1, mock_img2]

    with patch("PIL.Image.open") as mock_open:
        # Mock Pillow Image objects
        mock_pillow_img1 = MagicMock(spec=Image.Image)
        mock_pillow_img1.mode = "RGB"
        mock_pillow_img1.__enter__.return_value = mock_pillow_img1

        mock_pillow_img2 = MagicMock(spec=Image.Image)
        mock_pillow_img2.mode = "RGBA"
        mock_pillow_img2.__enter__.return_value = mock_pillow_img2

        # When convert is called, return another mock
        mock_pillow_img2_converted = MagicMock(spec=Image.Image)
        mock_pillow_img2_converted.mode = "RGB"
        mock_pillow_img2.convert.return_value = mock_pillow_img2_converted
        # Since it's used in the same context if converted?
        # Actually the code does: pillow_img = pillow_img.convert("RGB")
        # Then pillow_img.save is called.

        mock_open.side_effect = [mock_pillow_img1, mock_pillow_img2]

        media_info = extractor._extract_media(mock_ws, "Sheet1")

        assert len(media_info) == 2
        assert media_info[0]["coord"] == "A1"
        assert media_info[1]["coord"] == "B2"

        # Verify save was called on first image
        expected_path1 = tmp_path / "media" / "Sheet1_img_A1_0.png"
        mock_pillow_img1.save.assert_called_with(expected_path1, "PNG")

        # Verify RGBA to RGB conversion and save on second image
        mock_pillow_img2.convert.assert_called_with("RGB")
        expected_path2 = tmp_path / "media" / "Sheet1_img_B2_1.png"
        mock_pillow_img2_converted.save.assert_called_with(expected_path2, "PNG")

def test_extract_media_failure_skips_image(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Mock worksheet
    mock_ws = MagicMock()

    # Mock images: one valid, one invalid
    mock_img_ok = MagicMock()
    mock_img_ok.anchor._from.row = 0
    mock_img_ok.anchor._from.col = 0
    mock_img_ok.ref.read.return_value = b"good_data"

    mock_img_bad = MagicMock()
    mock_img_bad.anchor._from.row = 1
    mock_img_bad.anchor._from.col = 1
    mock_img_bad.ref.read.return_value = b"bad_data"

    mock_ws._images = [mock_img_ok, mock_img_bad]

    with patch("PIL.Image.open") as mock_open:
        # Mock Pillow Image objects
        mock_pillow_img_ok = MagicMock(spec=Image.Image)
        mock_pillow_img_ok.mode = "RGB"
        mock_pillow_img_ok.__enter__.return_value = mock_pillow_img_ok

        # Image.open raises exception for the second image
        mock_open.side_effect = [mock_pillow_img_ok, UnidentifiedImageError("Unknown format")]

        media_info = extractor._extract_media(mock_ws, "Sheet1")

        # Only the first image should be in media_info
        assert len(media_info) == 1
        assert media_info[0]["coord"] == "A1"
        assert media_info[0]["filename"] == "Sheet1_img_A1_0.png"

        # Verify that the good image save was called
        expected_path_ok = tmp_path / "media" / "Sheet1_img_A1_0.png"
        mock_pillow_img_ok.save.assert_called_with(expected_path_ok, "PNG")

def test_extract_media_value_error(tmp_path):
    """ValueError during Image.open should be caught and skipped."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()
    mock_img = MagicMock()
    mock_img.anchor._from.row = 0
    mock_img.anchor._from.col = 0
    mock_img.ref.read.return_value = b"some_data"
    mock_ws._images = [mock_img]

    with patch("PIL.Image.open") as mock_open:
        mock_open.side_effect = ValueError("Invalid image parameters")

        media_info = extractor._extract_media(mock_ws, "Sheet1")

        assert media_info == []

def test_extract_media_no_images(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()
    # Remove _images attribute if it exists
    if hasattr(mock_ws, "_images"):
        del mock_ws._images

    media_info = extractor._extract_media(mock_ws, "Sheet1")
    assert media_info == []
