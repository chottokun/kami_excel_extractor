import pytest
import subprocess
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.converter import ExcelConverter

def test_convert_input_file_not_found(output_dir):
    converter = ExcelConverter(output_dir)
    with pytest.raises(FileNotFoundError, match="Input file not found"):
        converter.convert(Path("non_existent.xlsx"))

@patch("subprocess.run")
def test_convert_libreoffice_failure(mock_run, output_dir, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy")

    mock_run.return_value = MagicMock(returncode=1, stderr="LibreOffice error")

    converter = ExcelConverter(output_dir)
    with pytest.raises(RuntimeError, match="LibreOffice conversion failed: LibreOffice error"):
        converter.convert(input_file)

@patch("subprocess.run")
def test_convert_pdftocairo_failure(mock_run, output_dir, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy")

    # First call (soffice) success
    res_pdf = MagicMock(returncode=0)
    # Second call (pdftocairo) failure
    res_png = MagicMock(returncode=1, stderr="pdftocairo error")
    mock_run.side_effect = [res_pdf, res_png]

    # We need the intermediate PDF to exist for the check
    pdf_path = output_dir / "test.pdf"
    pdf_path.write_text("fake pdf content")

    converter = ExcelConverter(output_dir)
    with pytest.raises(RuntimeError, match="pdftocairo conversion failed: pdftocairo error"):
        converter.convert(input_file)

    # PDF should still exist (or maybe not, but the code only unlinks on success)
    # Actually, the 'finally' block handles profile cleanup, but not PDF cleanup on pdftocairo failure?
    # Let's check the code:
    # if original_pdf.exists():
    #     original_pdf.unlink()
    # is AFTER the check for res_png.returncode != 0.
    # So it won't be unlinked if pdftocairo fails.
    assert pdf_path.exists()

@patch("subprocess.run")
def test_convert_success(mock_run, output_dir, tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.write_text("dummy")

    # Both calls success
    res_pdf = MagicMock(returncode=0)
    res_png = MagicMock(returncode=0)
    mock_run.side_effect = [res_pdf, res_png]

    # We need the intermediate PDF to exist for the check
    pdf_path = output_dir / "test.pdf"
    pdf_path.write_text("fake pdf content")

    # The output PNG is expected by the caller, though converter.convert just returns the path
    png_path = output_dir / "test.png"

    converter = ExcelConverter(output_dir)
    result = converter.convert(input_file)

    assert result == png_path
    assert not pdf_path.exists() # Should be unlinked on success
