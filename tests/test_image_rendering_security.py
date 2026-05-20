import html
from pathlib import Path
import pytest
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_block_level_image_rendering(doc_gen):
    """Verify that a line containing only an image is rendered as a div, not wrapped in a p."""
    md = "![alt](path/to/img.png)"
    html_out = doc_gen._simple_md_to_html(md)

    # Check that it's NOT wrapped in <p>
    assert "<p><div" not in html_out
    assert '<div class="image-container">' in html_out
    assert '<img src="path/to/img.png" alt="alt">' in html_out

def test_render_image_element_failure_escaping(doc_gen):
    """Verify that _render_image_element escapes its input even when parsing fails."""
    # Invalid image syntax that should trigger failure/fallback
    payload = "![unclosed(path"
    rendered = doc_gen._render_image_element(payload)

    assert rendered == html.escape(payload, quote=True)
    assert "(" in rendered # Should be escaped if needed, but here we just check it's not raw if it contained tags

def test_xss_in_block_image(doc_gen):
    """Verify XSS protection in block-level image tags."""
    payload_alt = '"><script>alert(1)</script>'
    payload_path = '"><script>alert(2)</script>'
    md = f"![{payload_alt}]({payload_path})"

    html_out = doc_gen._simple_md_to_html(md)

    assert "<script>" not in html_out
    assert html.escape(payload_alt) in html_out
    assert html.escape(payload_path, quote=True) in html_out
