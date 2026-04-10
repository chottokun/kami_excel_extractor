import pytest
import html
from unittest.mock import patch, MagicMock
from pathlib import Path
import tempfile
from kami_excel_extractor.document_generator import DocumentGenerator
from kami_excel_extractor.converter import ExcelConverter

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

def test_excel_converter_uses_timeout(tmp_path):
    input_file = tmp_path / "test.xlsx"
    input_file.touch()
    output_dir = tmp_path / "output"
    output_dir.mkdir()

    converter = ExcelConverter(output_dir)

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0

        def side_effect(*args, **kwargs):
            # Step 1 creates PDF
            if "soffice" in args[0][0]:
                pdf_file = output_dir / "test.pdf"
                pdf_file.touch()
            return mock_run.return_value

        mock_run.side_effect = side_effect

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

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_html_injection_in_text(doc_gen):
    md = "<script>alert('xss')</script>"
    html_out = doc_gen._simple_md_to_html(md)
    expected = html.escape("<script>alert('xss')</script>")
    assert f"<p>{expected}</p>" in html_out
    assert "<script>" not in html_out

def test_html_injection_in_header(doc_gen):
    md = "# <img src=x onerror=alert(1)>"
    html_out = doc_gen._simple_md_to_html(md)
    expected = html.escape("<img src=x onerror=alert(1)>")
    assert f"<h1>{expected}</h1>" in html_out
    assert "<img src=x" not in html_out

def test_html_injection_in_list(doc_gen):
    md = "- <svg onload=alert(1)>"
    html_out = doc_gen._simple_md_to_html(md)
    expected = html.escape("<svg onload=alert(1)>")
    assert f"<li>{expected}</li>" in html_out
    assert "<svg" not in html_out

def test_html_injection_in_table(doc_gen):
    md = "| <script> | alert(1) |\n| --- | --- |\n| </td> | </tr> |"
    html_out = doc_gen._simple_md_to_html(md)
    assert f"<th>{html.escape('<script>')}</th>" in html_out
    assert f"<td>{html.escape('</td>')}</td>" in html_out
    assert f"<td>{html.escape('</tr>')}</td>" in html_out

def test_html_injection_in_image_path(doc_gen):
    md = '![alt](some_path)'
    html_out = doc_gen._simple_md_to_html(md)
    assert 'src="some_path"' in html_out

    md2 = '![alt](" onerror="alert(1)")'
    html_out2 = doc_gen._simple_md_to_html(md2)
    assert 'src="&quot; onerror=&quot;alert(1"' in html_out2

def test_inline_styles_with_html(doc_gen):
    md = "**bold <script>**"
    html_out = doc_gen._simple_md_to_html(md)
    # We want <b>bold &lt;script&gt;</b>
    assert f"<b>bold {html.escape('<script>')}</b>" in html_out
