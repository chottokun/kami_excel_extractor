import pytest
import subprocess
import shutil
import asyncio
from pathlib import Path
from unittest.mock import MagicMock, patch
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_run_soffice_conversion_no_soffice(doc_gen, tmp_path):
    """soffice が見つからない場合のテスト"""
    with patch("shutil.which", return_value=None):
        temp_html = tmp_path / "test.html"
        temp_html.write_text("<html></html>")
        result = doc_gen._run_soffice_conversion(tmp_path, temp_html)
        assert result is None

def test_run_soffice_conversion_timeout(doc_gen, tmp_path):
    """soffice 変換がタイムアウトした場合のテスト"""
    with patch("shutil.which", return_value="/usr/bin/soffice"), \
         patch("kami_excel_extractor.document_generator.subprocess.run") as mock_run:

        mock_run.side_effect = subprocess.TimeoutExpired(cmd="soffice", timeout=60)

        temp_html = tmp_path / "test.html"
        temp_html.write_text("<html></html>")

        result = doc_gen._run_soffice_conversion(tmp_path, temp_html)
        assert result is None

@pytest.mark.asyncio
async def test_agenerate_pdf_basic(doc_gen, tmp_path):
    """非同期 PDF 生成の基本テスト"""
    with patch("kami_excel_extractor.document_generator.subprocess.run") as mock_run:
        def create_pdf_side_effect(cmd, **kwargs):
            outdir = Path(cmd[5])
            html_file = Path(cmd[6])
            expected_pdf = outdir / f"{html_file.stem}.pdf"
            expected_pdf.write_text("dummy pdf")
            return MagicMock(returncode=0)

        mock_run.side_effect = create_pdf_side_effect
        with patch("shutil.which", return_value="/usr/bin/soffice"):
            result = await doc_gen.agenerate_pdf("# Test", "test_report")
            assert result == tmp_path / "test_report.pdf"
            assert (tmp_path / "test_report.pdf").exists()

@pytest.mark.asyncio
async def test_agenerate_pdf_with_images(doc_gen, tmp_path):
    """画像を含む非同期 PDF 生成のテスト"""
    media_dir = tmp_path / "media"
    media_dir.mkdir()
    img_file = media_dir / "test.png"
    img_file.write_text("dummy image")

    md_content = "# Title\n![img](media/test.png)"

    with patch("kami_excel_extractor.document_generator.subprocess.run") as mock_run:
        def create_pdf_side_effect(cmd, **kwargs):
            outdir = Path(cmd[5])
            html_file = Path(cmd[6])
            expected_pdf = outdir / f"{html_file.stem}.pdf"
            expected_pdf.write_text("dummy pdf")
            return MagicMock(returncode=0)

        mock_run.side_effect = create_pdf_side_effect
        with patch("shutil.which", return_value="/usr/bin/soffice"):
            result = await doc_gen.agenerate_pdf(md_content, "test_img_report")
            assert result == tmp_path / "test_img_report.pdf"
            assert (tmp_path / "test_img_report.pdf").exists()

@pytest.mark.asyncio
async def test_agenerate_pdf_failure(doc_gen, tmp_path):
    """非同期 PDF 生成失敗のテスト"""
    with patch("kami_excel_extractor.document_generator.subprocess.run") as mock_run:
        mock_run.return_value = MagicMock(returncode=1)
        with patch("shutil.which", return_value="/usr/bin/soffice"):
            result = await doc_gen.agenerate_pdf("# Test", "test_fail")
            assert result is None

def test_parse_balanced_image_edge_cases(doc_gen):
    """_parse_balanced_image のエッジケーステスト"""
    # 括弧が閉じられていない
    assert doc_gen._parse_balanced_image("![alt](path") is None
    # alt の閉じカッコがない
    assert doc_gen._parse_balanced_image("![alt(path)") is None
    # ] の直後が ( でない
    assert doc_gen._parse_balanced_image("![alt] (path)") is None

@pytest.mark.asyncio
async def test_aresolve_single_image_not_found(doc_gen, tmp_path):
    """非同期画像解決で見つからない場合のテスト"""
    line = "![alt](missing.png)"
    result = await doc_gen._aresolve_single_image(line, [tmp_path], tmp_path / "tmp")
    assert result == line

def test_render_image_element_manual_call(doc_gen):
    """直接呼ばれないが残っている _render_image_element のテスト"""
    # 正常系
    assert '<img src="p.png" alt="a">' in doc_gen._render_image_element("![a](p.png)")
    # 不完全な形式
    assert doc_gen._render_image_element("![a]p.png") == "![a]p.png"
    assert doc_gen._render_image_element("![a](p.png") == "![a](p.png"
    # ] の直後が ( でない
    assert doc_gen._render_image_element("![alt] (path)") == "![alt] (path)"
    # 例外系 (あまり発生しないが coverage のため)
    with patch("html.escape", side_effect=Exception("mock error")):
        assert doc_gen._render_image_element("![a](p.png)") == "![a](p.png)"

def test_doc_gen_del_exception(tmp_path):
    """__del__ で例外が発生した場合のテスト (coverage のため)"""
    dg = DocumentGenerator(output_dir=tmp_path)
    with patch.object(dg._executor, "shutdown", side_effect=Exception("shutdown error")):
        dg.__del__() # Should not raise
