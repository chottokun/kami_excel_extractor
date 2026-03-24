import pytest
from unittest.mock import patch, MagicMock
from pathlib import Path
import tempfile
from kami_excel_extractor.document_generator import DocumentGenerator

def test_generate_pdf_uses_secure_temp_dir(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    generator = DocumentGenerator(output_dir)
    
    # 実際の一時ディレクトリ作成をフックして、パスを取得するが
    # 中でのファイル書き込みが成功するようにする
    with tempfile.TemporaryDirectory(prefix="pdf_gen_test_") as real_temp:
        with patch("tempfile.TemporaryDirectory") as mock_temp:
            mock_temp.return_value.__enter__.return_value = real_temp
            
            with patch("subprocess.run") as mock_run:
                mock_run.return_value.returncode = 0
                with patch("shutil.move"):
                    with patch("pathlib.Path.rglob") as mock_rglob:
                        mock_rglob.return_value = [Path(real_temp) / "test_report.pdf"]
                        # この呼び出しで tempfile.TemporaryDirectory が使われるはず
                        generator.generate_pdf("# Test", "test_report")
            
            mock_temp.assert_called_once()
            assert mock_temp.call_args[1].get("prefix") == "pdf_gen_"

from kami_excel_extractor.converter import ExcelConverter

def test_excel_converter_uses_timeout(tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    output_dir = tmp_path / "output"
    output_dir.mkdir()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0
        # Mock the PDF file creation so it doesn't fail the check
        pdf_file = output_dir / "test.pdf"
        pdf_file.touch()

        converter.convert(input_file)

        # Check if both subprocess calls had timeout
        assert mock_run.call_count == 2

        # First call: LibreOffice
        args, kwargs = mock_run.call_args_list[0]
        assert "soffice" in args[0]
        assert "timeout" in kwargs

        # Second call: pdftocairo
        args, kwargs = mock_run.call_args_list[1]
        assert "pdftocairo" in args[0]
        assert "timeout" in kwargs
