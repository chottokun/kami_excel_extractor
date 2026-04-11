import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path
from kami_excel_extractor.converter import ExcelConverter

@pytest.fixture(autouse=True)
def mock_shutil_which():
    with patch("shutil.which") as mock:
        mock.side_effect = lambda x: f"/usr/bin/{x}"
        yield mock

def test_convert_success(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test.pdf"
    png_file = output_dir / "test.png"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        mock_res.returncode = 0
        if "soffice" in args[0]:
            pdf_file.touch()
        return mock_res

    with patch("subprocess.run", side_effect=mock_run) as mock_subprocess:
        result = converter.convert(input_file)

        assert result == png_file
        assert not pdf_file.exists()
        assert mock_subprocess.call_count == 2
        # Verify absolute paths used
        assert mock_subprocess.call_args_list[0][0][0][0] == "/usr/bin/soffice"
        assert mock_subprocess.call_args_list[1][0][0][0] == "/usr/bin/pdftocairo"

def test_convert_input_missing(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "non_existent.xlsx"

    with pytest.raises(FileNotFoundError, match="Input file not found"):
        converter.convert(input_file)

def test_convert_soffice_failure(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 1
        mock_run.return_value.stderr = "LibreOffice Error"

        with pytest.raises(RuntimeError, match="LibreOffice conversion failed: LibreOffice Error"):
            converter.convert(input_file)

def test_convert_pdf_missing(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0
        # No PDF file created

        with pytest.raises(FileNotFoundError, match="PDF not found after conversion"):
            converter.convert(input_file)

def test_convert_pdftocairo_failure(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test.pdf"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        if "/usr/bin/soffice" in args[0]:
            mock_res.returncode = 0
            pdf_file.touch()
        elif "/usr/bin/pdftocairo" in args[0]:
            mock_res.returncode = 1
            mock_res.stderr = "pdftocairo Error"
        return mock_res

    with patch("subprocess.run", side_effect=mock_run):
        with pytest.raises(RuntimeError, match="pdftocairo conversion failed: pdftocairo Error"):
            converter.convert(input_file)

        # Intermediate PDF should be unlinked even if pdftocairo fails
        assert not pdf_file.exists()
