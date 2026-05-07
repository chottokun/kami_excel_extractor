import pytest
import logging
import io
from pathlib import Path
from unittest.mock import MagicMock, patch
from PIL import Image
from kami_excel_extractor.extractor import MetadataExtractor, MAX_IMAGE_BYTES

def test_extract_media_mock_read_exception(tmp_path, caplog):
    """Test when img.ref is a Mock and read() raises an exception."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    # Create a custom mock class that has "Mock" in its name
    class CustomMockImageRef(MagicMock):
        pass

    mock_img = MagicMock()
    mock_img.ref = CustomMockImageRef()
    mock_img.ref.read.side_effect = IOError("Mock read failure")
    mock_img.anchor._from.row = 0
    mock_img.anchor._from.col = 0

    mock_ws._images = [mock_img]

    with caplog.at_level(logging.WARNING):
        media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 1
    assert media_info[0]["error"] == "unidentified_format"
    assert "Failed to extract image at A1 on sheet Sheet1: Mock read failure" in caplog.text

def test_extract_media_chunked_read_exception(tmp_path, caplog):
    """Test when chunked reading via read(chunk_size) raises an exception."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    class FaultyStream:
        def read(self, size=-1):
            raise IOError("Chunked read failure")

    mock_img = MagicMock()
    mock_img.ref = FaultyStream()
    mock_img.anchor._from.row = 1
    mock_img.anchor._from.col = 1

    mock_ws._images = [mock_img]

    with caplog.at_level(logging.WARNING):
        media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 1
    assert media_info[0]["error"] == "unidentified_format"
    assert "Failed to extract image at B2 on sheet Sheet1: Chunked read failure" in caplog.text

def test_extract_media_chunked_read_limit_exceeded(tmp_path, caplog):
    """Test when chunked reading exceeds MAX_IMAGE_BYTES."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    class LargeStream:
        def __init__(self):
            self.count = 0
        def read(self, size):
            self.count += 1
            if self.count > 10: # Just to prevent infinite loop if code were broken
                return b""
            return b"a" * size

    mock_img = MagicMock()
    mock_img.ref = LargeStream()
    mock_img.anchor._from.row = 2
    mock_img.anchor._from.col = 2

    mock_ws._images = [mock_img]

    # We need to make sure the loop hits the limit.
    with patch("kami_excel_extractor.extractor.MAX_IMAGE_BYTES", 100):
        with caplog.at_level(logging.WARNING):
            media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 0 # It should continue and not add to media_info if it breaks due to limit
    assert "Skipping large image at C3 on Sheet1 (stream exceeds limit)" in caplog.text

def test_extract_media_no_readable_attribute(tmp_path, caplog):
    """Test when img.ref has no readable data attribute."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    mock_img = MagicMock()
    mock_img.ref = object() # Simple object with no read/getvalue/getbuffer
    mock_img.anchor._from.row = 3
    mock_img.anchor._from.col = 3

    mock_ws._images = [mock_img]

    with caplog.at_level(logging.WARNING):
        media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 1
    assert media_info[0]["error"] == "unidentified_format"
    assert "Image reference has no readable data attribute" in caplog.text

def test_extract_media_bytes_too_large(tmp_path, caplog):
    """Test when img.ref is bytes but exceeds MAX_IMAGE_BYTES."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    mock_img = MagicMock()
    mock_img.ref = b"too large"
    mock_img.anchor._from.row = 4
    mock_img.anchor._from.col = 4

    mock_ws._images = [mock_img]

    with patch("kami_excel_extractor.extractor.MAX_IMAGE_BYTES", 5):
        with caplog.at_level(logging.WARNING):
            media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 0
    assert "Skipping large image at E5 on Sheet1" in caplog.text

def test_extract_media_getbuffer_success(tmp_path):
    """Test when img.ref has getbuffer (like io.BytesIO) and no read."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    # Create a small valid PNG in memory
    img = Image.new('RGB', (10, 10), color = 'red')
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')

    class BufferOnly:
        def __init__(self, data):
            self.data = data
        def getbuffer(self):
            return self.data.getbuffer()

    mock_img = MagicMock()
    mock_img.ref = BufferOnly(img_byte_arr)
    mock_img.anchor._from.row = 5
    mock_img.anchor._from.col = 5

    mock_ws._images = [mock_img]

    media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 1
    assert media_info[0]["filename"] is not None
    assert (tmp_path / "media" / media_info[0]["filename"]).exists()

def test_extract_media_getvalue_success(tmp_path):
    """Test when img.ref has getvalue and no read/getbuffer."""
    extractor = MetadataExtractor(output_dir=tmp_path)
    mock_ws = MagicMock()

    # Create a small valid PNG in memory
    img = Image.new('RGB', (10, 10), color = 'green')
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='PNG')

    class ValueOnly:
        def __init__(self, data):
            self.data = data
        def getvalue(self):
            return self.data.getvalue()

    mock_img = MagicMock()
    mock_img.ref = ValueOnly(img_byte_arr)
    mock_img.anchor._from.row = 6
    mock_img.anchor._from.col = 6

    mock_ws._images = [mock_img]

    media_info = extractor._extract_media(mock_ws, "Sheet1")

    assert len(media_info) == 1
    assert media_info[0]["filename"] is not None
    assert (tmp_path / "media" / media_info[0]["filename"]).exists()
