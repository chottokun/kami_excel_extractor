import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.extractor import MetadataExtractor
from kami_excel_extractor.schema import ExtractionOptions
import asyncio

@pytest.fixture
def realistic_xlsx():
    return Path("tests/assets/realistic_business_report.xlsx")

@pytest.mark.asyncio
async def test_realistic_visual_intelligence_pipeline(tmp_path, realistic_xlsx):
    """実戦的なデータを用いて、図表データの抽出とHTML注入が正常に行われることを検証"""
    extractor = KamiExcelExtractor(output_dir=tmp_path)
    
    # テスト環境に依存する部分（画像変換やLLM呼び出し）をモック
    with patch("litellm.acompletion") as mock_completion, \
         patch("kami_excel_extractor.core.ExcelConverter") as mock_conv_cls, \
         patch("kami_excel_extractor.extractor.MetadataExtractor.extract") as mock_extract:
        
        # 1. Extractorのモック結果を定義
        # 画像が E3 (Summaryシート) と A3, C3 (SitePhotosシート) にあるとする
        mock_extract.return_value = {
            "sheets": {
                "Summary": {
                    "html": "<table><tr><td data-coord=\"E3\">Graph Area</td></tr></table>",
                    "media": [{"coord": "E3", "filename": "Summary_img_E3_0.png", "type": "image"}],
                    "media_map": {"E3": [{"coord": "E3", "filename": "Summary_img_E3_0.png", "type": "image"}]}
                }
            }
        }
        
        # 2. 画像ファイルが存在するように見せかける
        (tmp_path / "media").mkdir(parents=True, exist_ok=True)
        (tmp_path / "media" / "Summary_img_E3_0.png").touch()
        
        # 3. Converterのモック
        mock_conv = mock_conv_cls.return_value
        mock_conv.convert.return_value = tmp_path / "full_page.png"
        (tmp_path / "full_page.png").touch()

        # 4. LLMレスポンスのモック
        # 1つ目: 要約, 2つ目: 図表データ抽出
        media_data_content = "[図表データ]\n| 月 | 実績 |\n|---|---|\n| 2024-01 | 1100 |"
        media_resp = MagicMock()
        media_resp.choices[0].message.content = media_data_content
        
        final_resp = MagicMock()
        final_resp.choices[0].message.content = '```json\n{"summary_data": "ok"}\n```'
        
        mock_completion.side_effect = [media_resp, media_resp, final_resp]
        
        # 実行
        options = ExtractionOptions(include_visual_summaries=True, use_visual_context=True)
        result = await extractor.aextract_structured_data(realistic_xlsx, options=options)
        
        # 検証
        assert "Summary" in result["sheets"]
        
        # LLMへの最終プロンプトに注入されたデータが含まれているか
        sheet_call_args = mock_completion.call_args_list[-1]
        messages = sheet_call_args[1]["messages"]
        user_content = next(m["content"] for m in messages if m["role"] == "user")
        html_text = next(c["text"] for c in user_content if c["type"] == "text" and "データソース" in c["text"])
        
        assert "[図表データ(E3)]" in html_text
        assert "| 月 | 実績 |" in html_text
        assert "2024-01" in html_text

def test_extractor_with_realistic_anchors(tmp_path, realistic_xlsx):
    """実戦的なエクセルファイルのアンカー情報を正しく抽出できるか検証"""
    me = MetadataExtractor(output_dir=tmp_path)
    result = me.extract(realistic_xlsx)
    
    # SummaryシートのE3に画像があること
    summary_media = result["sheets"]["Summary"]["media"]
    assert any(m["coord"] == "E3" for m in summary_media)
    
    # SitePhotosシートに画像があること
    photo_media = result["sheets"]["SitePhotos"]["media"]
    assert len(photo_media) >= 1
    assert any(m["coord"] in ["A3", "C3"] for m in photo_media)
