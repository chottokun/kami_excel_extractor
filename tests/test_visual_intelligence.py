import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.extractor import MetadataExtractor
from kami_excel_extractor.schema import ExtractionOptions
import asyncio
import io

@pytest.fixture
def mock_excel_with_image(tmp_path):
    import openpyxl
    from openpyxl.drawing.image import Image as OpenpyxlImage
    from PIL import Image as PILImage
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "グラフタイトル"
    
    # 有効なPNG画像の作成
    img_path = tmp_path / "dummy_chart.png"
    PILImage.new('RGB', (100, 100), color='red').save(img_path, format='PNG')
    
    # Excelに画像を配置 (文字列アンカーを使用)
    img = OpenpyxlImage(str(img_path))
    img.anchor = "A2"
    ws.add_image(img)
    
    excel_path = tmp_path / "chart_test.xlsx"
    wb.save(excel_path)
    return excel_path

@pytest.mark.asyncio
async def test_visual_data_injection_pipeline(tmp_path, mock_excel_with_image):
    """画像(グラフ)からデータが抽出され、HTMLに注入されるパイプラインを検証"""
    extractor = KamiExcelExtractor(output_dir=tmp_path)
    
    with patch("litellm.acompletion") as mock_completion, \
         patch("kami_excel_extractor.core.ExcelConverter") as mock_conv_cls:
        
        mock_conv = mock_conv_cls.return_value
        mock_conv.convert.return_value = tmp_path / "dummy_page.png"
        (tmp_path / "dummy_page.png").touch()

        # LLMレスポンスのモック
        # 1. 要約, 2. 図表データ抽出
        media_resp = MagicMock()
        media_resp.choices[0].message.content = "[図表データ]\n| 年月 | 売上 |\n|---|---|\n| 2023-01 | 100 |"
        
        # 3. シート解析
        final_resp = MagicMock()
        final_resp.choices[0].message.content = '```json\n{"data": "success"}\n```'
        
        mock_completion.side_effect = [media_resp, media_resp, final_resp]
        
        options = ExtractionOptions(include_visual_summaries=True, use_visual_context=True)
        result = await extractor.aextract_structured_data(mock_excel_with_image, options=options)
        
        # メッセージの検証
        sheet_call_args = mock_completion.call_args_list[-1]
        messages = sheet_call_args[1]["messages"]
        user_content = next(m["content"] for m in messages if m["role"] == "user")
        html_text = next(c["text"] for c in user_content if c["type"] == "text" and "データソース" in c["text"])
        
        assert "[座標 A2 の図表データ]" in html_text
        assert "| 年月 | 売上 |" in html_text

def test_media_map_coordinate_linking(tmp_path, mock_excel_with_image):
    """Extractorが座標とメディアを正しく紐付けているか検証"""
    from kami_excel_extractor.extractor import MetadataExtractor
    me = MetadataExtractor(output_dir=tmp_path)
    result = me.extract(mock_excel_with_image)
    
    sheet_info = result["sheets"]["Sheet"]
    assert "media_map" in sheet_info
    assert "A2" in sheet_info["media_map"]
