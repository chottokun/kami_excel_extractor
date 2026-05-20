import html
from pathlib import Path

import pytest

from kami_excel_extractor.core import KamiExcelExtractor


@pytest.fixture
def extractor(tmp_path):
    return KamiExcelExtractor(output_dir=tmp_path)


def test_visual_insights_xss_protection(extractor):
    """Verify that visual_data injected into HTML is properly escaped."""
    coord = "A1"
    payload = "<script>alert('xss')</script>"
    items = [{"visual_data": payload}]

    insight_html = extractor._format_visual_insights(coord, items)

    assert "<script>" not in insight_html
    assert html.escape(payload) in insight_html
    assert f"[図表データ({coord})]" in insight_html


def test_visual_insights_injection(extractor):
    """Verify that insights are correctly injected and escaped in the final HTML."""
    html_content = '<table><tr><td data-coord="A1">Cell</td></tr></table>'
    payload = "<b>Injection</b>"
    media_map = {"A1": [{"visual_data": payload}]}

    final_html = extractor._inject_visual_data_to_html(html_content, media_map)

    assert 'data-coord="A1"' in final_html
    assert html.escape(payload) in final_html
    assert "<b>" not in final_html
