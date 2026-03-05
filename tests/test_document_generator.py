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

@patch("subprocess.run")
def test_generate_pdf_success(mock_run, doc_gen, tmp_path):
    # subprocess.run のモック
    mock_run.return_value = MagicMock(returncode=0)
    
    # 実際には soffice は PDF を生成しないので、手動で作成してシミュレートする
    # DocumentGenerator.generate_pdf は tmp_dir 内で PDF を探し、output_dir に移動する
    # そのため、少し工夫が必要。generate_pdf の内部で作成される tmp_dir を特定しにくいので、
    # shutil.move をモックするか、あるいは glob の結果を操作する。
    
    with patch("shutil.move") as mock_move, \
         patch("pathlib.Path.glob") as mock_glob:
        
        mock_pdf = MagicMock(spec=Path)
        mock_glob.return_value = [mock_pdf]
        
        result = doc_gen.generate_pdf("# Test Content", "test_report")
        
        assert mock_run.called
        assert result == tmp_path / "test_report.pdf"
        assert mock_move.called

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
