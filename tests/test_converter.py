import pytest
import subprocess
import shutil
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.converter import ExcelConverter

@pytest.fixture
def converter(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return ExcelConverter(output_dir)

def test_convert_input_file_not_found(converter, tmp_path):
    input_file = tmp_path / "non_existent.xlsx"
    with pytest.raises(FileNotFoundError, match="Input file not found"):
        converter.convert(input_file)

@patch("kami_excel_extractor.converter.subprocess.run")
@patch("kami_excel_extractor.converter.shutil.rmtree")
def test_convert_libreoffice_failure(mock_rmtree, mock_run, converter, tmp_path):
    # Create input file
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy content")

    # Mock subprocess failure for LibreOffice
    mock_run.return_value = MagicMock(returncode=1, stderr="Some LibreOffice error")

    with pytest.raises(RuntimeError, match="LibreOffice conversion failed: Some LibreOffice error"):
        converter.convert(input_file)

    # Verify it was called with soffice
    args, _ = mock_run.call_args
    assert "soffice" in args[0]

@patch("kami_excel_extractor.converter.subprocess.run")
@patch("kami_excel_extractor.converter.shutil.rmtree")
def test_convert_pdf_not_found(mock_rmtree, mock_run, converter, tmp_path):
    # Create input file
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy content")

    # Mock subprocess success for LibreOffice but don't create PDF
    mock_run.return_value = MagicMock(returncode=0)

    with pytest.raises(FileNotFoundError, match="PDF not found after conversion"):
        converter.convert(input_file)

@patch("kami_excel_extractor.converter.subprocess.run")
@patch("kami_excel_extractor.converter.shutil.rmtree")
def test_convert_pdftocairo_failure(mock_rmtree, mock_run, converter, tmp_path):
    # Create input file
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy content")

    # Create PDF file to pass Step 1 check
    pdf_file = converter.output_dir / "test.pdf"
    pdf_file.write_text("dummy pdf")

    # Mock subprocess: 1st call (soffice) success, 2nd call (pdftocairo) failure
    mock_run.side_effect = [
        MagicMock(returncode=0),
        MagicMock(returncode=1, stderr="pdftocairo error")
    ]

    with pytest.raises(RuntimeError, match="pdftocairo conversion failed: pdftocairo error"):
        converter.convert(input_file)

@patch("kami_excel_extractor.converter.subprocess.run")
@patch("kami_excel_extractor.converter.shutil.rmtree")
def test_convert_success(mock_rmtree, mock_run, converter, tmp_path):
    # Create input file
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy content")

    # Create PDF file to pass Step 1 check
    pdf_file = converter.output_dir / "test.pdf"
    pdf_file.write_text("dummy pdf")

    # Create PNG file to simulate pdftocairo success
    png_file = converter.output_dir / "test.png"
    png_file.write_text("dummy png")

    # Mock subprocess success for both calls
    mock_run.return_value = MagicMock(returncode=0)

    result = converter.convert(input_file)

    assert result == png_file
    assert not pdf_file.exists()  # intermediate PDF should be unlinked
    assert mock_rmtree.called
