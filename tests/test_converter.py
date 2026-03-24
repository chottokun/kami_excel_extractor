import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path
from kami_excel_extractor.converter import ExcelConverter

def test_convert_success(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        # Mock soffice run
        mock_soffice_res = MagicMock()
        mock_soffice_res.returncode = 0

        # Mock pdftocairo run
        mock_pdftocairo_res = MagicMock()
        mock_pdftocairo_res.returncode = 0

        mock_run.side_effect = [mock_soffice_res, mock_pdftocairo_res]

        # We need to simulate the creation of the PDF file by soffice
        pdf_file = output_dir / "test.pdf"
        pdf_file.touch()

        # We also need to simulate the creation of the PNG file by pdftocairo
        png_file = output_dir / "test.png"
        png_file.touch()

        result = converter.convert(input_file)

        assert result == png_file
        assert not pdf_file.exists()  # Should be unlinked
        assert mock_run.call_count == 2

def test_convert_input_file_not_found(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    input_file = tmp_path / "non_existent.xlsx"

    converter = ExcelConverter(output_dir)

    with pytest.raises(FileNotFoundError, match="Input file not found"):
        converter.convert(input_file)

def test_convert_soffice_failure(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        mock_res = MagicMock()
        mock_res.returncode = 1
        mock_res.stderr = "LibreOffice Error"
        mock_run.return_value = mock_res

        with pytest.raises(RuntimeError, match="LibreOffice conversion failed"):
            converter.convert(input_file)

def test_convert_pdf_not_found(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        mock_res = MagicMock()
        mock_res.returncode = 0
        mock_run.return_value = mock_res

        # No PDF file created

        with pytest.raises(FileNotFoundError, match="PDF not found after conversion"):
            converter.convert(input_file)

def test_convert_pdftocairo_failure(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        # Mock soffice run
        mock_soffice_res = MagicMock()
        mock_soffice_res.returncode = 0

        # Mock pdftocairo run
        mock_pdftocairo_res = MagicMock()
        mock_pdftocairo_res.returncode = 1
        mock_pdftocairo_res.stderr = "pdftocairo Error"

        mock_run.side_effect = [mock_soffice_res, mock_pdftocairo_res]

        pdf_file = output_dir / "test.pdf"
        pdf_file.touch()

        with pytest.raises(RuntimeError, match="pdftocairo conversion failed"):
            converter.convert(input_file)
