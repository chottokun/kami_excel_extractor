import html
from pathlib import Path
from kami_excel_extractor.document_generator import DocumentGenerator

def test_image_source_escaping():
    doc_gen = DocumentGenerator(output_dir=Path("/tmp"))
    # Parentheses in path to test regex robustness, and XSS payload to test escaping
    malicious_path = 'image(1).png"> <script>alert("xss")</script>'
    md = f'![alt]({malicious_path})'
    html_out = doc_gen._simple_md_to_html(md)

    expected_escaped_path = html.escape(malicious_path, quote=True)
    assert f'src="{expected_escaped_path}"' in html_out
    assert '<script>' not in html_out

def test_image_alt_escaping():
    doc_gen = DocumentGenerator(output_dir=Path("/tmp"))
    malicious_alt = '"> <script>alert("xss")</script>'
    md = f'![{malicious_alt}](path/to/img.png)'
    html_out = doc_gen._simple_md_to_html(md)

    expected_escaped_alt = html.escape(malicious_alt, quote=True)
    assert f'alt="{expected_escaped_alt}"' in html_out
    assert '<script>' not in html_out
