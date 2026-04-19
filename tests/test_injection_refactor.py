import pytest
from kami_excel_extractor.core import KamiExcelExtractor

def test_inject_visual_data_to_html_baseline(output_dir):
    extractor = KamiExcelExtractor(output_dir=output_dir)

    html_content = """
    <table>
        <tr>
            <td data-coord="A1">Value A1</td>
            <td data-coord='B2'>Value B2</td>
            <td data-coord="C3">Value C3</td>
        </tr>
    </table>
    """

    media_map = {
        "A1": [
            {"visual_data": "Insight for A1", "filename": "a1.png"}
        ],
        "B2": [
            {"visual_data": "Insight 1 for B2"},
            {"visual_data": "Insight 2 for B2"}
        ],
        "C3": [
            {"other_field": "no visual data"}
        ],
        "D4": [
            {"visual_data": "Insight for D4"}
        ]
    }

    result = extractor._inject_visual_data_to_html(html_content, media_map)

    # Check A1 (double quotes)
    assert 'data-coord="A1" <div class=\'visual-insight\'>[図表データ(A1)]: Insight for A1</div>' in result

    # Check B2 (single quotes)
    assert "data-coord='B2' <div class='visual-insight'>[図表データ(B2)]: Insight 1 for B2</div>\n<div class='visual-insight'>[図表データ(B2)]: Insight 2 for B2</div>" in result

    # Check C3 (no visual_data)
    assert 'data-coord="C3">Value C3' in result
    assert "[図表データ(C3)]" not in result

    # Check D4 (not in HTML)
    assert "[図表データ(D4)]" not in result

def test_inject_visual_data_to_html_multiple_occurrences(output_dir):
    extractor = KamiExcelExtractor(output_dir=output_dir)

    html_content = '<td data-coord="A1"></td><td data-coord="A1"></td>'
    media_map = {
        "A1": [{"visual_data": "Insight"}]
    }

    result = extractor._inject_visual_data_to_html(html_content, media_map)

    # str.replace replaces all occurrences
    assert result.count("visual-insight") == 2
