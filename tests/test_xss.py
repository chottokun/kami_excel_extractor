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
    # The regex RE_IMAGE supports one level of balanced parentheses.
    # If we have unbalanced parentheses or multiple levels, it might stop early or fail to match.
    md = '![alt](img.png" onerror="alert(1)")'
    html_out = doc_gen._simple_md_to_html(md)
    assert 'onerror="alert(1)"' not in html_out
    # With balanced parentheses support, "img.png\" onerror=\"alert(1)\"" is matched as the path.
    assert 'src="img.png&quot; onerror=&quot;alert(1)&quot;"' in html_out

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

def test_image_path_with_parentheses(doc_gen):
    # Test that image paths with parentheses are correctly handled and escaped
    md = "![alt](image(1).png)"
    html_out = doc_gen._simple_md_to_html(md)
    assert 'src="image(1).png"' in html_out
    assert 'alt="alt"' in html_out

    # Mixed case with potential XSS and parentheses
    md = '![alt" onerror="alert(1)](image(2).png?x=y&z=w)'
    html_out = doc_gen._simple_md_to_html(md)
    assert 'src="image(2).png?x=y&amp;z=w"' in html_out
    assert 'alt="alt&quot; onerror=&quot;alert(1)"' in html_out
    assert 'onerror="alert(1)"' not in html_out

def test_multiple_images_on_same_line(doc_gen):
    # Test that the regex is not too greedy and handles multiple images on the same line
    # However, _simple_md_to_html currently processes images line by line if they start with ![
    # If there are multiple images on one line, and it doesn't start with ![, it might be treated as paragraph.
    # If it starts with ![, only the first image is matched by search() and then the loop moves to next line.

    md = "![img1](path1.png) ![img2](path2.png)"
    html_out = doc_gen._simple_md_to_html(md)
    # Current implementation of _simple_md_to_html only handles one image per line if it starts with ![
    assert 'src="path1.png"' in html_out
    # Since search() is used on the whole line, and it's not a global replace in _simple_md_to_html
    # Only the first match is rendered by _render_image_element.

    # Let's check how it handles image followed by text in parentheses
    md = "![img](path.png) (additional info)"
    html_out = doc_gen._simple_md_to_html(md)
    assert 'src="path.png"' in html_out
    assert '(additional info)' not in html_out # Wait, _render_image_element only returns the <div><img></div>
    # The rest of the line is ignored in the current stripped.startswith('![') block.
