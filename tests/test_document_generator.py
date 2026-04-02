import pytest
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_simple_md_to_html_basic(doc_gen):
    md = "# Title\n- Item 1\n- Item 2\nParagraph text."
    html = doc_gen._simple_md_to_html(md)
    
    assert "<h1>Title</h1>" in html
    assert "<ul>" in html
    assert "<li>Item 1</li>" in html
    assert "<li>Item 2</li>" in html
    assert "<p>Paragraph text.</p>" in html
    assert 'lang="ja"' in html

def test_simple_md_to_html_table(doc_gen):
    md = "| Header 1 | Header 2 |\n| --- | --- |\n| Cell 1 | Cell 2 |"
    html = doc_gen._simple_md_to_html(md)
    
    assert "<table>" in html
    assert "<th>Header 1</th>" in html
    assert "<td>Cell 1</td>" in html

def test_simple_md_to_html_visual_summary(doc_gen):
    md = "[画像概要] これはテストです。"
    html = doc_gen._simple_md_to_html(md)
    assert '<div class="visual-summary">' in html
    assert 'これはテストです。' in html

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_discovery_normal(mock_run, doc_gen, tmp_path):
    """標準的なパス（expected_pdf が存在する）で PDF が生成・移動されることを確認"""
    def create_pdf_side_effect(cmd, **kwargs):
        # cmd: ["soffice", ..., "--outdir", tmp_dir, temp_html]
        outdir = Path(cmd[5])
        html_file = Path(cmd[6])
        expected_pdf = outdir / f"{html_file.stem}.pdf"
        expected_pdf.write_text("dummy pdf content")
        return MagicMock(returncode=0)

    mock_run.side_effect = create_pdf_side_effect
    
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    
    assert result == tmp_path / "test_report.pdf"
    assert (tmp_path / "test_report.pdf").exists()
    mock_run.assert_called_once()

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_discovery_fallback(mock_run, doc_gen, tmp_path):
    """expected_pdf が存在せず、fallback (rglob) で PDF が見つかるケースを確認"""
    def create_differently_named_pdf_side_effect(cmd, **kwargs):
        outdir = Path(cmd[5])
        # 期待される名前とは異なる名前で PDF を作成
        unexpected_pdf = outdir / "different_name.pdf"
        unexpected_pdf.write_text("dummy pdf content")
        return MagicMock(returncode=0)

    mock_run.side_effect = create_differently_named_pdf_side_effect

    result = doc_gen.generate_pdf("# Test Content", "test_report")

    assert result == tmp_path / "test_report.pdf"
    assert (tmp_path / "test_report.pdf").exists()
    mock_run.assert_called_once()

@patch("subprocess.run")
def test_generate_pdf_subprocess_error(mock_run, doc_gen):
    """soffice がエラー（非ゼロ終了）を返した場合のテスト"""
    mock_run.return_value = MagicMock(returncode=1)
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    assert result is None

@patch("subprocess.run")
def test_generate_pdf_exception(mock_run, doc_gen):
    """subprocess.run が例外を投げた場合のテスト"""
    mock_run.side_effect = OSError("Subprocess crash")
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    assert result is None

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_unexpected_exception(mock_run, doc_gen):
    """予期せぬ例外（Exception）が発生した場合のテスト"""
    mock_run.side_effect = RuntimeError("Unexpected crash")
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    assert result is None

@patch("subprocess.run")
def test_generate_pdf_no_output(mock_run, doc_gen):
    """soffice は成功したが、PDFファイルが生成されなかった場合のテスト"""
    mock_run.return_value = MagicMock(returncode=0)
    with patch("pathlib.Path.rglob") as mock_rglob:
        mock_rglob.return_value = []
        result = doc_gen.generate_pdf("# Test Content", "test_report")
        assert result is None

def test_resolve_images_to_tmpdir(doc_gen, tmp_path):
    # 画像ファイルを作成
    media_dir = tmp_path / "media"
    media_dir.mkdir()
    img_file = media_dir / "test.png"
    img_file.write_text("dummy binary content")
    
    md_content = "![画像](media/test.png)"
    work_dir = tmp_path / "work"
    work_dir.mkdir()
    
    resolved_md = doc_gen._resolve_images_to_tmpdir(md_content, work_dir)
    
    # パスが書き換えられているか確認
    assert "file://" in resolved_md
    assert (work_dir / "test.png").exists()

