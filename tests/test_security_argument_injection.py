import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path
from kami_excel_extractor.converter import ExcelConverter
from kami_excel_extractor.document_generator import DocumentGenerator

def test_excel_converter_uses_absolute_paths(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    # Use a relative path to ensure the fix actually makes it absolute
    input_file = Path("test.xlsx")
    # We need to touch it so exists() passes
    input_file.touch()

    try:
        def mock_run_side_effect(args, **kwargs):
            if any("soffice" in str(arg) for arg in args):
                pdf_file = output_dir.resolve() / "test.pdf"
                pdf_file.touch()
            return MagicMock(returncode=0)

        with patch("subprocess.run", side_effect=mock_run_side_effect) as mock_run, \
             patch("shutil.which", side_effect=lambda x: f"/usr/bin/{x}"):

            converter.convert(input_file)

            # First call: soffice
            # [ "soffice", "-env:UserInstallation=...", "--headless", "--convert-to", "pdf", "--outdir", str(self.output_dir), str(input_file) ]
            soffice_args = mock_run.call_args_list[0][0][0]
            outdir = soffice_args[6]
            input_path = soffice_args[7]
            assert Path(outdir).is_absolute(), f"Expected absolute path for outdir, got {outdir}"
            assert Path(input_path).is_absolute(), f"Expected absolute path for input, got {input_path}"

            # Second call: pdftocairo
            # [ "pdftocairo", "-png", "-singlefile", str(original_pdf), str(self.output_dir / input_file.stem) ]
            pdftocairo_args = mock_run.call_args_list[1][0][0]
            pdf_path = pdftocairo_args[3]
            output_base = pdftocairo_args[4]
            assert Path(pdf_path).is_absolute(), f"Expected absolute path for pdf_path, got {pdf_path}"
            assert Path(output_base).is_absolute(), f"Expected absolute path for output_base, got {output_base}"
    finally:
        if input_file.exists():
            input_file.unlink()

def test_document_generator_uses_absolute_paths(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    generator = DocumentGenerator(output_dir)

    with patch("subprocess.run") as mock_run, \
         patch("shutil.which", side_effect=lambda x: f"/usr/bin/{x}"):
        mock_run.return_value.returncode = 0

        # We need to mock rglob because generate_pdf uses it to find the pdf
        with patch("pathlib.Path.rglob") as mock_rglob:
            mock_rglob.return_value = [tmp_path / "test.pdf"]
            with patch("shutil.move"):
                generator.generate_pdf("# Test", "test_report")

        # [ "soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(temp_html) ]
        soffice_args = mock_run.call_args[0][0]
        outdir = soffice_args[5]
        html_path = soffice_args[6]
        assert Path(outdir).is_absolute(), f"Expected absolute path for outdir, got {outdir}"
        assert Path(html_path).is_absolute(), f"Expected absolute path for html_path, got {html_path}"
