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
