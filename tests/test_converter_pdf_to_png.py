import pytest
import subprocess
from unittest.mock import patch, MagicMock
from pathlib import Path
from kami_excel_extractor.converter import ExcelConverter

@pytest.fixture
def converter(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return ExcelConverter(output_dir)

@pytest.fixture
def pdf_path(tmp_path):
    path = tmp_path / "test.pdf"
    path.touch()
    return path

@pytest.fixture
def png_path(tmp_path):
    return tmp_path / "test.png"

@pytest.fixture
def mock_run():
    with patch("subprocess.run") as mock:
        yield mock

@pytest.fixture
def mock_which():
    with patch("shutil.which") as mock:
        yield mock

def test_try_pdftocairo_success(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "pdftocairo" else None

    def mock_run_impl(args, **kwargs):
        png_path.touch()
        return MagicMock(returncode=0)

    mock_run.side_effect = mock_run_impl

    assert converter._try_pdftocairo(pdf_path, png_path) is True
    assert png_path.exists()

    # Check arguments
    args = mock_run.call_args[0][0]
    assert args[0] == "/usr/bin/pdftocairo"
    assert args[1] == "-png"
    assert args[2] == "-singlefile"
    assert args[3] == str(pdf_path.resolve())
    assert args[4] == str(png_path.with_suffix("").resolve())

def test_try_pdftocairo_not_found(converter, pdf_path, png_path, mock_which):
    mock_which.return_value = None
    assert converter._try_pdftocairo(pdf_path, png_path) is False

def test_try_pdftocairo_failure(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.return_value = "/usr/bin/pdftocairo"
    mock_run.return_value = MagicMock(returncode=1, stderr="error")
    assert converter._try_pdftocairo(pdf_path, png_path) is False

def test_try_pdftocairo_exception(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.return_value = "/usr/bin/pdftocairo"
    mock_run.side_effect = OSError("failed to execute")
    assert converter._try_pdftocairo(pdf_path, png_path) is False

def test_try_fitz_success(converter, pdf_path, png_path):
    mock_fitz = MagicMock()
    mock_doc = mock_fitz.open.return_value
    mock_page = mock_doc.load_page.return_value
    mock_pix = mock_page.get_pixmap.return_value

    def mock_save(path):
        png_path.touch()
    mock_pix.save.side_effect = mock_save

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        assert converter._try_fitz(pdf_path, png_path) is True
        assert png_path.exists()
        mock_fitz.open.assert_called_once_with(str(pdf_path.resolve()))
        mock_pix.save.assert_called_once_with(str(png_path.resolve()))

def test_try_fitz_import_error(converter, pdf_path, png_path):
    with patch.dict("sys.modules", {"fitz": None}):
        assert converter._try_fitz(pdf_path, png_path) is False

def test_try_fitz_exception(converter, pdf_path, png_path):
    mock_fitz = MagicMock()
    mock_fitz.open.side_effect = Exception("Fitz error")
    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        assert converter._try_fitz(pdf_path, png_path) is False

def test_try_imagemagick_magick_success(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "magick" else None

    def mock_run_impl(args, **kwargs):
        png_path.touch()
        return MagicMock(returncode=0)

    mock_run.side_effect = mock_run_impl

    assert converter._try_imagemagick(pdf_path, png_path) is True
    assert png_path.exists()

    args = mock_run.call_args[0][0]
    assert args[0] == "/usr/bin/magick"
    assert args[1] == "-density"
    assert args[2] == str(converter.dpi)
    assert args[3] == f"{pdf_path.resolve()}[0]"
    assert args[4] == str(png_path.resolve())

def test_try_imagemagick_convert_success(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "convert" else None

    def mock_run_impl(args, **kwargs):
        png_path.touch()
        return MagicMock(returncode=0)

    mock_run.side_effect = mock_run_impl

    assert converter._try_imagemagick(pdf_path, png_path) is True
    assert png_path.exists()

    args = mock_run.call_args[0][0]
    assert args[0] == "/usr/bin/convert"

def test_try_imagemagick_not_found(converter, pdf_path, png_path, mock_which):
    mock_which.return_value = None
    assert converter._try_imagemagick(pdf_path, png_path) is False

def test_try_imagemagick_failure(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.return_value = "/usr/bin/magick"
    mock_run.return_value = MagicMock(returncode=1)
    assert converter._try_imagemagick(pdf_path, png_path) is False

def test_try_imagemagick_exception(converter, pdf_path, png_path, mock_which, mock_run):
    mock_which.return_value = "/usr/bin/magick"
    mock_run.side_effect = subprocess.SubprocessError("timeout")
    assert converter._try_imagemagick(pdf_path, png_path) is False

def test_convert_pdf_to_png_pdftocairo_primary(converter, pdf_path, png_path):
    with patch.object(converter, "_try_pdftocairo", return_value=True) as m1, \
         patch.object(converter, "_try_fitz") as m2, \
         patch.object(converter, "_try_imagemagick") as m3:

        converter._convert_pdf_to_png(pdf_path, png_path)
        m1.assert_called_once()
        m2.assert_not_called()
        m3.assert_not_called()

def test_convert_pdf_to_png_fallback_to_fitz(converter, pdf_path, png_path):
    with patch.object(converter, "_try_pdftocairo", return_value=False) as m1, \
         patch.object(converter, "_try_fitz", return_value=True) as m2, \
         patch.object(converter, "_try_imagemagick") as m3:

        converter._convert_pdf_to_png(pdf_path, png_path)
        m1.assert_called_once()
        m2.assert_called_once()
        m3.assert_not_called()

def test_convert_pdf_to_png_fallback_to_imagemagick(converter, pdf_path, png_path):
    with patch.object(converter, "_try_pdftocairo", return_value=False) as m1, \
         patch.object(converter, "_try_fitz", return_value=False) as m2, \
         patch.object(converter, "_try_imagemagick", return_value=True) as m3:

        converter._convert_pdf_to_png(pdf_path, png_path)
        m1.assert_called_once()
        m2.assert_called_once()
        m3.assert_called_once()

def test_convert_pdf_to_png_all_fail(converter, pdf_path, png_path):
    with patch.object(converter, "_try_pdftocairo", return_value=False), \
         patch.object(converter, "_try_fitz", return_value=False), \
         patch.object(converter, "_try_imagemagick", return_value=False):

        with pytest.raises(RuntimeError, match="All PDF to PNG conversion methods failed"):
            converter._convert_pdf_to_png(pdf_path, png_path)

def test_try_fitz_multi_success(converter, pdf_path, tmp_path):
    output_prefix = converter.output_dir / "test_prefix"
    mock_fitz = MagicMock()
    mock_doc = mock_fitz.open.return_value
    mock_doc.__len__.return_value = 2
    mock_page = mock_doc.load_page.return_value
    mock_pix = mock_page.get_pixmap.return_value

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        result = converter._try_fitz_multi(pdf_path, output_prefix)

        assert len(result) == 2
        assert result[0] == converter.output_dir / "test_prefix-1.png"
        assert result[1] == converter.output_dir / "test_prefix-2.png"

        assert mock_fitz.open.call_count == 1
        assert mock_doc.load_page.call_count == 2
        assert mock_pix.save.call_count == 2
        mock_doc.close.assert_called_once()

def test_try_fitz_multi_exception(converter, pdf_path):
    output_prefix = converter.output_dir / "test_prefix"
    mock_fitz = MagicMock()
    mock_fitz.open.side_effect = Exception("Fitz open error")

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        with pytest.raises(RuntimeError, match="Multi-page PDF to PNG conversion failed"):
            converter._try_fitz_multi(pdf_path, output_prefix)

def test_try_fitz_multi_no_pages(converter, pdf_path):
    output_prefix = converter.output_dir / "test_prefix"
    mock_fitz = MagicMock()
    mock_doc = mock_fitz.open.return_value
    mock_doc.__len__.return_value = 0

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        with pytest.raises(RuntimeError, match="Multi-page PDF to PNG conversion failed"):
            converter._try_fitz_multi(pdf_path, output_prefix)
