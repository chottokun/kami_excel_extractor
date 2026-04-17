import pytest
from pathlib import Path
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_xss_injection(doc_gen):
    # Test Bold XSS
    md = "**<script>alert(1)</script>**"
    html_out = doc_gen._simple_md_to_html(md)
    # < and > should be escaped
    assert "<script>" not in html_out
    assert "<b>&lt;script&gt;alert(1)&lt;/script&gt;</b>" in html_out

    # Test Image Alt XSS - Breakout attempt
    # If quote=True is working, the " will be escaped and won't break out of alt attribute
    md = '![alt" onerror="alert(1)](img.png)'
    html_out = doc_gen._simple_md_to_html(md)
    # It should NOT have a raw onerror attribute on the img tag
    assert 'onerror="alert(1)"' not in html_out
    # It SHOULD have the escaped version in the alt attribute
    # Note: html.escape by default does NOT escape " unless quote=True is passed.
    # We want to ensure quote=True is used.
    assert 'alt="alt&quot; onerror=&quot;alert(1)"' in html_out

    # Test Image URL XSS - Breakout attempt
    # The regex RE_IMAGE = re.compile(r'!\[(.*?)\]\((.*?)\)') will stop at the first ')'
    md = '![alt](img.png" onerror="alert(1)")'
    html_out = doc_gen._simple_md_to_html(md)
    assert 'onerror="alert(1)"' not in html_out
    # img_path will be: img.png" onerror="alert(1"
    assert 'src="img.png&quot; onerror=&quot;alert(1"' in html_out

    # Test Table XSS
    md = "| <script>alert(1)</script> |"
    html_out = doc_gen._simple_md_to_html(md)
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in html_out
    assert "<script>" not in html_out

    # Test Header XSS
    md = "# <script>alert(1)</script>"
    html_out = doc_gen._simple_md_to_html(md)
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in html_out
    assert "<script>" not in html_out

    # Test Visual Summary XSS
    md = "[画像概要] <script>alert(1)</script>"
    html_out = doc_gen._simple_md_to_html(md)
    assert "&lt;script&gt;alert(1)&lt;/script&gt;" in html_out
    assert '<div class="visual-summary">[画像概要] &lt;script&gt;alert(1)&lt;/script&gt;</div>' in html_out
