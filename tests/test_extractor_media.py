import pytest
from unittest.mock import MagicMock
from pathlib import Path
from kami_excel_extractor.extractor import MetadataExtractor

def test_extract_media_no_images(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    # Create a mock that definitely does not have _images
    ws = MagicMock()
    del ws._images

    result = extractor._extract_media(ws, "Sheet1")
    assert result == []

def test_extract_media_with_images(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    ws = MagicMock()

    # Mock image 1
    img1 = MagicMock()
    img1.anchor._from.row = 0
    img1.anchor._from.col = 0
    img1.ref.read.return_value = b"fake image data 1"

    # Mock image 2 (using getvalue)
    img2 = MagicMock()
    img2.anchor._from.row = 5
    img2.anchor._from.col = 2
    # Ensure img2.ref does not have 'read' to trigger 'getvalue' branch
    del img2.ref.read
    img2.ref.getvalue.return_value = b"fake image data 2"

    ws._images = [img1, img2]

    result = extractor._extract_media(ws, "Sheet1")

    assert len(result) == 2

    # Verify image 1
    assert result[0] == {"coord": "A1", "filename": "Sheet1_img_A1_0.png", "type": "image"}
    save_path1 = tmp_path / "media" / "Sheet1_img_A1_0.png"
    assert save_path1.exists()
    assert save_path1.read_bytes() == b"fake image data 1"

    # Verify image 2
    assert result[1] == {"coord": "C6", "filename": "Sheet1_img_C6_1.png", "type": "image"}
    save_path2 = tmp_path / "media" / "Sheet1_img_C6_1.png"
    assert save_path2.exists()
    assert save_path2.read_bytes() == b"fake image data 2"

def test_extract_media_exception_handling(tmp_path):
    extractor = MetadataExtractor(output_dir=tmp_path)
    ws = MagicMock()

    # Mock image 1 (fails on read)
    img1 = MagicMock()
    img1.anchor._from.row = 0
    img1.anchor._from.col = 0
    img1.ref.read.side_effect = Exception("Read error")

    # Mock image 2 (succeeds)
    img2 = MagicMock()
    img2.anchor._from.row = 1
    img2.anchor._from.col = 1
    img2.ref.read.return_value = b"good data"

    ws._images = [img1, img2]

    # Should not raise exception due to try-except block in _extract_media
    result = extractor._extract_media(ws, "Sheet1")

    # Only image 2 should be in results
    assert len(result) == 1
    assert result[0]["coord"] == "B2"

    # Verify file for image 2 exists
    assert (tmp_path / "media" / "Sheet1_img_B2_1.png").exists()
