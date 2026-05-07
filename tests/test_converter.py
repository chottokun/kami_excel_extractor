import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path
from kami_excel_extractor.converter import ExcelConverter

@pytest.fixture(autouse=True)
def mock_shutil_which():
    with patch("shutil.which") as mock:
        mock.side_effect = lambda x: f"/usr/bin/{x}"
        yield mock

@pytest.fixture(autouse=True)
def mock_uuid():
    with patch("uuid.uuid4") as mock:
        mock_val = MagicMock()
        mock_val.__str__.return_value = "12345678-1234-1234-1234-123456789012"
        mock.return_value = mock_val
        yield mock

def test_convert_success(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test_12345678.pdf"
    png_file = output_dir / "test_12345678.png"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        mock_res.returncode = 0
        if "/usr/bin/soffice" in args[0]:
            pdf_file.touch()
        if "/usr/bin/pdftocairo" in args[0]:
            png_file.touch()
        return mock_res

    with patch("subprocess.run", side_effect=mock_run) as mock_subprocess:
        result = converter.convert(input_file)

        assert result == png_file
        assert not pdf_file.exists()
        assert mock_subprocess.call_count == 2
        # Verify absolute paths used
        assert mock_subprocess.call_args_list[0][0][0][0] == "/usr/bin/soffice"
        assert mock_subprocess.call_args_list[1][0][0][0] == "/usr/bin/pdftocairo"

        # Verify pdftocairo arguments are resolved
        pdftocairo_args = mock_subprocess.call_args_list[1][0][0]
        assert Path(pdftocairo_args[3]).is_absolute()
        assert Path(pdftocairo_args[4]).is_absolute()

def test_convert_dpi_propagation(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir, dpi=300)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test_12345678.pdf"
    png_file = output_dir / "test_12345678.png"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        mock_res.returncode = 0
        if "/usr/bin/soffice" in args[0]:
            pdf_file.touch()
        elif "/usr/bin/magick" in args[0]:
            png_file.touch()
        return mock_res

    # Mock pdftocairo as missing to trigger ImageMagick
    with patch("shutil.which", side_effect=lambda x: f"/usr/bin/{x}" if x != "pdftocairo" else None):
        with patch("subprocess.run", side_effect=mock_run) as mock_subprocess:
            with patch.dict("sys.modules", {"fitz": None}):
                result = converter.convert(input_file)
                assert result == png_file

                # Find the magick call
                magick_call = next(c for c in mock_subprocess.call_args_list if "/usr/bin/magick" in c[0][0])
                args = magick_call[0][0]
                assert args[2] == "300" # DPI
                assert Path(args[3].replace("[0]", "")).is_absolute()
                assert Path(args[4]).is_absolute()

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

def test_convert_fallback_to_fitz(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test_12345678.pdf"
    png_file = output_dir / "test_12345678.png"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        if "/usr/bin/soffice" in args[0]:
            mock_res.returncode = 0
            pdf_file.touch()
        elif "/usr/bin/pdftocairo" in args[0]:
            mock_res.returncode = 1
            mock_res.stderr = "pdftocairo Error"
        return mock_res

    mock_fitz = MagicMock()
    mock_doc = mock_fitz.open.return_value
    mock_page = mock_doc.load_page.return_value

    def mock_save(path):
        png_file.touch()

    mock_page.get_pixmap.return_value.save.side_effect = mock_save

    with patch("subprocess.run", side_effect=mock_run):
        with patch.dict("sys.modules", {"fitz": mock_fitz}):
            result = converter.convert(input_file)
            assert result == png_file
            assert png_file.exists()
            assert not pdf_file.exists()
            mock_fitz.open.assert_called_once()

def test_convert_fallback_to_imagemagick(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test_12345678.pdf"
    png_file = output_dir / "test_12345678.png"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        if "/usr/bin/soffice" in args[0]:
            mock_res.returncode = 0
            pdf_file.touch()
        elif "/usr/bin/pdftocairo" in args[0]:
            mock_res.returncode = 1
        elif "/usr/bin/magick" in args[0]:
            mock_res.returncode = 0
            png_file.touch()
        return mock_res

    # Mock fitz as missing
    with patch("subprocess.run", side_effect=mock_run):
        with patch.dict("sys.modules", {"fitz": None}):
            result = converter.convert(input_file)
            assert result == png_file
            assert png_file.exists()
            assert not pdf_file.exists()

def test_convert_all_fallbacks_fail(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    pdf_file = output_dir / "test_12345678.pdf"

    def mock_run(args, **kwargs):
        mock_res = MagicMock()
        if "/usr/bin/soffice" in args[0]:
            mock_res.returncode = 0
            pdf_file.touch()
        else:
            mock_res.returncode = 1
            mock_res.stderr = "All fail"
        return mock_res

    with patch("subprocess.run", side_effect=mock_run):
        with patch.dict("sys.modules", {"fitz": None}):
            with pytest.raises(RuntimeError, match="All PDF to PNG conversion methods failed"):
                converter.convert(input_file)

        # Intermediate PDF should be unlinked even if all fail
        assert not pdf_file.exists()
