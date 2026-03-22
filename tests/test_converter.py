import pytest
import subprocess
from pathlib import Path
from unittest.mock import MagicMock, patch
import shutil
from kami_excel_extractor.converter import ExcelConverter

@pytest.fixture
def converter(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return ExcelConverter(output_dir=output_dir)

def test_convert_success(converter, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy excel content")

    output_png = converter.output_dir / "test.png"
    output_pdf = converter.output_dir / "test.pdf"

    def side_effect(args, **kwargs):
        if "soffice" in args:
            output_pdf.write_text("dummy pdf content")
            return MagicMock(returncode=0)
        elif "pdftocairo" in args:
            output_png.write_text("dummy png content")
            return MagicMock(returncode=0)
        return MagicMock(returncode=1)

    with patch("subprocess.run", side_effect=side_effect) as mock_run:
        result = converter.convert(input_file)

        assert result == output_png
        assert output_png.exists()
        assert not output_pdf.exists()  # Should be unlinked
        assert mock_run.call_count == 2

        # Verify soffice call
        soffice_call = mock_run.call_args_list[0]
        args = soffice_call[0][0]
        assert "soffice" in args
        assert "--headless" in args
        assert "--convert-to" in args
        assert "pdf" in args
        assert str(converter.output_dir) in args
        assert str(input_file) in args

        # Verify pdftocairo call
        pdftocairo_call = mock_run.call_args_list[1]
        args = pdftocairo_call[0][0]
        assert "pdftocairo" in args
        assert "-png" in args
        assert str(output_pdf) in args

def test_convert_input_not_found(converter, tmp_path):
    input_file = tmp_path / "non_existent.xlsx"

    with pytest.raises(FileNotFoundError, match="Input file not found"):
        converter.convert(input_file)

def test_convert_soffice_failure(converter, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy excel content")

    with patch("subprocess.run") as mock_run:
        mock_run.return_value = MagicMock(returncode=1, stderr="soffice error")

        with pytest.raises(RuntimeError, match="LibreOffice conversion failed: soffice error"):
            converter.convert(input_file)

def test_convert_pdf_not_found(converter, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy excel content")

    with patch("subprocess.run") as mock_run:
        mock_run.return_value = MagicMock(returncode=0)
        # We don't create the PDF file here

        with pytest.raises(FileNotFoundError, match="PDF not found after conversion"):
            converter.convert(input_file)

def test_convert_pdftocairo_failure(converter, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy excel content")
    output_pdf = converter.output_dir / "test.pdf"

    def side_effect(args, **kwargs):
        if "soffice" in args:
            output_pdf.write_text("dummy pdf content")
            return MagicMock(returncode=0)
        elif "pdftocairo" in args:
            return MagicMock(returncode=1, stderr="pdftocairo error")
        return MagicMock(returncode=1)

    with patch("subprocess.run", side_effect=side_effect):
        with pytest.raises(RuntimeError, match="pdftocairo conversion failed: pdftocairo error"):
            converter.convert(input_file)

def test_convert_cleanup_profile(converter, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy excel content")

    # We want to check if shutil.rmtree was called with the correct path
    with patch("subprocess.run") as mock_run, \
         patch("shutil.rmtree") as mock_rmtree, \
         patch("uuid.uuid4") as mock_uuid:

        mock_uuid.return_value.hex = "test_uuid"
        mock_run.return_value = MagicMock(returncode=1, stderr="error") # Force failure

        with pytest.raises(RuntimeError):
            converter.convert(input_file)

        mock_rmtree.assert_called_once_with("/tmp/lo_profile_test_uuid", ignore_errors=True)
