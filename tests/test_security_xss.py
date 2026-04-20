import pytest
import html
from pathlib import Path
from kami_excel_extractor.document_generator import DocumentGenerator

@pytest.fixture
def doc_gen(tmp_path):
    return DocumentGenerator(output_dir=tmp_path)

def test_xss_in_header(doc_gen):
    payload = "<script>alert('header')</script>"
    md = f"# {payload}"
    html_out = doc_gen._simple_md_to_html(md)
    expected_escaped = html.escape(payload)
    assert f"<h1>{expected_escaped}</h1>" in html_out
    assert "<script>" not in html_out

def test_xss_in_list_item(doc_gen):
    payload = "<img src=x onerror=alert(1)>"
    md = f"- {payload}"
    html_out = doc_gen._simple_md_to_html(md)
    expected_escaped = html.escape(payload)
    assert f"<li>{expected_escaped}</li>" in html_out
    assert "<img" not in html_out

def test_xss_in_paragraph(doc_gen):
    payload = "<svg/onload=alert(1)>"
    md = payload
    html_out = doc_gen._simple_md_to_html(md)
    expected_escaped = html.escape(payload)
    assert f"<p>{expected_escaped}</p>" in html_out
    assert "<svg" not in html_out

def test_xss_in_table_cell(doc_gen):
    payload_h = "<iframe src='javascript:alert(1)'></iframe>"
    payload_d = "<a href='javascript:alert(1)'>Click me</a>"
    md = f"| {payload_h} |\n| --- |\n| {payload_d} |"
    html_out = doc_gen._simple_md_to_html(md)

    expected_h = html.escape(payload_h)
    expected_d = html.escape(payload_d)

    assert f"<th>{expected_h}</th>" in html_out
    assert f"<td>{expected_d}</td>" in html_out
    assert "<iframe>" not in html_out
    assert "<a " not in html_out

def test_xss_with_inline_styles(doc_gen):
    # Verify that bold styling works but XSS within it is escaped
    payload = "<script>alert(1)</script>"
    md = f"- **bold {payload}**"
    html_out = doc_gen._simple_md_to_html(md)

    # Expected: <li><b>bold &lt;script&gt;alert(1)&lt;/script&gt;</b></li>
    expected_inner = f"bold {html.escape(payload)}"
    assert f"<li><b>{expected_inner}</b></li>" in html_out
    assert "<script>" not in html_out

def test_xss_in_image(doc_gen):
    payload_alt = '"> <script>alert("alt")</script>'
    # Use a payload without parentheses in the path to avoid regex confusion
    payload_path = '"> <script>alert"path"</script>'
    md = f"![{payload_alt}]({payload_path})"
    html_out = doc_gen._simple_md_to_html(md)

    expected_alt = html.escape(payload_alt)
    expected_path = html.escape(payload_path, quote=True)

    assert f'alt="{expected_alt}"' in html_out
    assert f'src="{expected_path}"' in html_out
    assert "<script>" not in html_out
def test_xss_in_visual_summary(doc_gen):
    payload = "<img src=x onerror=alert(1)>"
    md = f"[画像概要] {payload}"
    html_out = doc_gen._simple_md_to_html(md)
    expected_escaped = html.escape(payload)
    assert '<div class="visual-summary">' in html_out
    assert expected_escaped in html_out
    assert "<img" not in html_out

def test_xss_with_quotes_in_inline_styles(doc_gen):
    # Verify that single and double quotes are escaped before bold styling
    payload = '"><script>alert(1)</script>'
    md = f"**{payload}**"
    html_out = doc_gen._simple_md_to_html(md)

    # html.escape(payload) should escape " and >
    expected_escaped = html.escape(payload)
    assert f"<b>{expected_escaped}</b>" in html_out
    assert "<script>" not in html_out
    assert '"><b>' not in html_out

def test_xss_bold_interaction(doc_gen):
    # Test that bold tags don't get double-escaped or broken by partial payloads
    # This specifically checks the interaction between escaping and RE_BOLD
    payload = "normal **bold** <script>"
    md = f"**{payload}**"
    html_out = doc_gen._simple_md_to_html(md)

    # The regex RE_BOLD (non-greedy) will match **normal ** and ** <script>**
    # but the key is that <script> is escaped.
    assert "<b>normal </b>" in html_out
    assert "&lt;script&gt;" in html_out
    assert "<script>" not in html_out
