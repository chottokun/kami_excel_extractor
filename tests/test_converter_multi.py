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
def output_prefix(tmp_path):
    return tmp_path / "output" / "test_prefix"

def test_try_fitz_multi_success(converter, pdf_path, output_prefix):
    mock_fitz = MagicMock()
    mock_doc = mock_fitz.open.return_value
    mock_doc.__len__.return_value = 2

    mock_page1 = MagicMock()
    mock_page2 = MagicMock()
    mock_doc.load_page.side_effect = [mock_page1, mock_page2]

    mock_pix1 = mock_page1.get_pixmap.return_value
    mock_pix2 = mock_page2.get_pixmap.return_value

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        pngs = converter._try_fitz_multi(pdf_path, output_prefix)

        assert len(pngs) == 2
        assert pngs[0] == converter.output_dir / f"{output_prefix.name}-1.png"
        assert pngs[1] == converter.output_dir / f"{output_prefix.name}-2.png"

        mock_fitz.open.assert_called_once_with(str(pdf_path.resolve()))
        assert mock_doc.load_page.call_count == 2
        assert mock_pix1.save.called
        assert mock_pix2.save.called
        mock_doc.close.assert_called_once()

def test_try_fitz_multi_exception(converter, pdf_path, output_prefix):
    mock_fitz = MagicMock()
    mock_fitz.open.side_effect = Exception("Fitz open error")

    with patch.dict("sys.modules", {"fitz": mock_fitz}):
        with pytest.raises(RuntimeError, match="Multi-page PDF to PNG conversion failed"):
            converter._try_fitz_multi(pdf_path, output_prefix)

def test_convert_pdf_to_multi_png_pdftocairo_success(converter, pdf_path, output_prefix):
    with patch("shutil.which") as mock_which, \
         patch("subprocess.run") as mock_run:

        mock_which.return_value = "/usr/bin/pdftocairo"
        mock_run.return_value = MagicMock(returncode=0)

        # Create dummy output files
        png1 = converter.output_dir / f"{output_prefix.name}-1.png"
        png2 = converter.output_dir / f"{output_prefix.name}-2.png"
        png1.touch()
        png2.touch()

        pngs = converter._convert_pdf_to_multi_png(pdf_path, output_prefix)

        assert len(pngs) == 2
        assert pngs[0] == png1
        assert pngs[1] == png2
        mock_run.assert_called_once()
        assert "pdftocairo" in mock_run.call_args[0][0][0]

def test_convert_pdf_to_multi_png_fallback_to_fitz(converter, pdf_path, output_prefix):
    with patch("shutil.which") as mock_which, \
         patch.object(converter, "_try_fitz_multi") as mock_try_fitz:

        # Scenario 1: pdftocairo missing
        mock_which.return_value = None
        mock_try_fitz.return_value = [Path("fake.png")]

        converter._convert_pdf_to_multi_png(pdf_path, output_prefix)
        mock_try_fitz.assert_called_once()

        mock_try_fitz.reset_mock()

        # Scenario 2: pdftocairo fails
        mock_which.return_value = "/usr/bin/pdftocairo"
        with patch("subprocess.run") as mock_run:
            mock_run.side_effect = Exception("pdftocairo error")
            converter._convert_pdf_to_multi_png(pdf_path, output_prefix)
            mock_try_fitz.assert_called_once()
