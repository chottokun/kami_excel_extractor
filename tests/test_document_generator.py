import pytest
import html
import subprocess
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_simple_md_to_html_basic(doc_gen):
    md = "# Title\n- Item 1\n- Item 2\nParagraph text."
    html_output = doc_gen._simple_md_to_html(md)
    
    assert "<h1>Title</h1>" in html_output
    assert "<ul>" in html_output
    assert "<li>Item 1</li>" in html_output
    assert "<li>Item 2</li>" in html_output
    assert "<p>Paragraph text.</p>" in html_output
    assert 'lang="ja"' in html_output

def test_simple_md_to_html_table(doc_gen):
    md = "| Header 1 | Header 2 |\n| --- | --- |\n| Cell 1 | Cell 2 |"
    html_output = doc_gen._simple_md_to_html(md)
    
    assert "<table>" in html_output
    assert "<th>Header 1</th>" in html_output
    assert "<td>Cell 1</td>" in html_output

def test_simple_md_to_html_visual_summary(doc_gen):
    md = "[画像概要] これはテストです。"
    html_output = doc_gen._simple_md_to_html(md)
    assert '<div class="visual-summary">' in html_output
    assert 'これはテストです。' in html_output

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_success(mock_run, doc_gen, tmp_path):
    # subprocess.run のモック
    mock_run.return_value = MagicMock(returncode=0)
    
    # 実際には soffice は PDF を生成しないので、手動で作成してシミュレートする
    # DocumentGenerator.generate_pdf は tmp_dir 内で PDF を探し、output_dir に移動する
    # そのため、少し工夫が必要。generate_pdf の内部で作成される tmp_dir を特定しにくいので、
    # shutil.move をモックするか、あるいは glob の結果を操作する。
    
    with patch("shutil.move") as mock_move, \
         patch("pathlib.Path.rglob") as mock_rglob:
        
        mock_pdf = MagicMock(spec=Path)
        mock_rglob.return_value = [mock_pdf]
        
        result = doc_gen.generate_pdf("# Test Content", "test_report")
        
        mock_run.assert_called_once()
        assert result == tmp_path / "test_report.pdf"
        assert mock_move.called

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_subprocess_error(mock_run, doc_gen):
    """soffice がエラー（非ゼロ終了）を返した場合のテスト"""
    mock_run.return_value = MagicMock(returncode=1)
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    assert result is None

@patch("kami_excel_extractor.document_generator.subprocess.run")
def test_generate_pdf_exception(mock_run, doc_gen):
    """subprocess.run が例外を投げた場合のテスト"""
    # 既存のコードが OSError や subprocess.SubprocessError をキャッチすることを考慮し、
    # generic Exception もキャッチされるため、テストとしては有効
    mock_run.side_effect = Exception("Subprocess crash")
    result = doc_gen.generate_pdf("# Test Content", "test_report")
    assert result is None

@patch("kami_excel_extractor.document_generator.subprocess.run")
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

def test_xss_protection_in_paragraph(doc_gen):
    """段落におけるXSS保護のテスト"""
    xss_payload = "<script>alert('xss')</script>"
    md = f"Normal text {xss_payload}"
    html_output = doc_gen._simple_md_to_html(md)

    # ペイロードがエスケープされていることを確認
    escaped_payload = html.escape(xss_payload)
    assert escaped_payload in html_output
    assert xss_payload not in html_output

def test_xss_protection_with_inline_styles(doc_gen):
    """インラインスタイル（太字）を伴うXSS保護のテスト"""
    xss_payload = "<script>alert('xss')</script>"
    md = f"**{xss_payload}**"
    html_output = doc_gen._simple_md_to_html(md)

    # エスケープされたペイロードが<b>タグ内にあることを確認
    escaped_payload = html.escape(xss_payload)
    assert f"<b>{escaped_payload}</b>" in html_output
    assert xss_payload not in html_output

def test_xss_protection_in_other_elements(doc_gen):
    """ヘッダー、リスト、テーブルにおけるXSS保護のテスト"""
    xss_payload = "<script>alert('xss')</script>"
    md = f"# {xss_payload}\n- {xss_payload}\n| {xss_payload} |\n| --- |\n| {xss_payload} |"
    html_output = doc_gen._simple_md_to_html(md)

    escaped_payload = html.escape(xss_payload)
    assert f"<h1>{escaped_payload}</h1>" in html_output
    assert f"<li>{escaped_payload}</li>" in html_output
    assert f"<th>{escaped_payload}</th>" in html_output
    assert f"<td>{escaped_payload}</td>" in html_output
