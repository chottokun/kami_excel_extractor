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
