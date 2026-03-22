import pytest
from unittest.mock import MagicMock
from pathlib import Path
from kami_excel_extractor.extractor import MetadataExtractor

def test_extract_media_handles_exception(tmp_path):
    """
    Test that _extract_media correctly handles exceptions during image data reading,
    skipping the failed image and continuing with others.
    """
    extractor = MetadataExtractor(output_dir=tmp_path)

    # Mock worksheet
    ws = MagicMock()

    # Mock image 1: Success (read() method)
    img1 = MagicMock()
    img1.anchor._from.row = 0
    img1.anchor._from.col = 0
    img1.ref = MagicMock()
    img1.ref.read.return_value = b"fake image data 1"
    # Ensure it has read but not getvalue for this one
    del img1.ref.getvalue

    # Mock image 2: Failure (raises Exception)
    img2 = MagicMock()
    img2.anchor._from.row = 1
    img2.anchor._from.col = 1
    img2.ref = MagicMock()
    img2.ref.read.side_effect = Exception("Read error")
    del img2.ref.getvalue

    # Mock image 3: Success (getvalue() method)
    img3 = MagicMock()
    img3.anchor._from.row = 2
    img3.anchor._from.col = 2
    img3.ref = MagicMock()
    del img3.ref.read
    img3.ref.getvalue.return_value = b"fake image data 3"

    ws._images = [img1, img2, img3]

    # Execute extraction
    media_info = extractor._extract_media(ws, "Sheet1")

    # Verify results: should have 2 successful images, image 2 should be skipped
    assert len(media_info) == 2

    # Image 1 verification
    assert media_info[0]["coord"] == "A1"
    assert media_info[0]["filename"] == "Sheet1_img_A1_0.png"
    save_path1 = tmp_path / "media" / media_info[0]["filename"]
    assert save_path1.exists()
    assert save_path1.read_bytes() == b"fake image data 1"

    # Image 3 verification (it became the 2nd item in media_info)
    assert media_info[1]["coord"] == "C3"
    assert media_info[1]["filename"] == "Sheet1_img_C3_2.png"
    save_path3 = tmp_path / "media" / media_info[1]["filename"]
    assert save_path3.exists()
    assert save_path3.read_bytes() == b"fake image data 3"

    # Verify that the directory for media exists
    assert (tmp_path / "media").is_dir()
