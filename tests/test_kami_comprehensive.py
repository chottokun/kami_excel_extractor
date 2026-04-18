import pytest
from pathlib import Path
from unittest.mock import patch, MagicMock
from kami_excel_extractor.extractor import MetadataExtractor
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

@pytest.fixture
def complex_xlsx():
    return Path("tests/assets/complex_kami_sample.xlsx")

def test_extractor_complex_layout(tmp_path, complex_xlsx):
    """複雑なレイアウトの抽出結果（HTML/スタイル）を詳細に検証"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    result = extractor.extract(complex_xlsx)
    
    sheet_data = result["sheets"]["ComplexLayout"]
    html = sheet_data["html"]
    
    # 1. ヘッダー (A1:F2)
    # 結合セルの属性確認
    assert 'colspan="6"' in html
    assert 'rowspan="2"' in html
    assert 'background-color: #CCE5FF' in html # 水色
    assert 'border-top: 3px solid black' in html # 太い線
    
    # 2. カテゴリ列 (A6:A8)
    assert 'rowspan="3"' in html
    assert 'background-color: #E0E0E0' in html # グレー
    
    # 3. 座標情報
    assert 'data-coord="A1"' in html
    assert 'data-coord="E4"' in html
    assert 'data-coord="B10"' in html

def test_core_prompt_generation_with_styles(tmp_path, complex_xlsx):
    """core.py が生成するプロンプトがスタイル情報を考慮しているか検証"""
    extractor = KamiExcelExtractor(output_dir=tmp_path)
    
    # _aextract_single_sheet 内でのメッセージ構築をフック
    with patch("litellm.acompletion") as mock_completion:
        mock_completion.return_value.choices[0].message.content = '{"data": "mocked"}'
        
        # 非同期実行
        import asyncio
        options = ExtractionOptions(use_visual_context=False)
        asyncio.run(extractor.aextract_structured_data(complex_xlsx, options=options))
        
        # 呼ばれた際のメッセージを確認
        # aextract_structured_data -> _aextract_single_sheet -> _build_sheet_messages
        args, kwargs = mock_completion.call_args
        messages = kwargs["messages"]
        
        user_msg = next(m["content"] for m in messages if m["role"] == "user")
        text_content = next(c["text"] for c in user_msg if c["type"] == "text" and "データソース" in c["text"])
        
        # スタイル属性や座標に関する説明が含まれているか
        assert "CSSスタイル属性(style)" in text_content
        assert "border属性" in text_content
        assert "data-coord属性" in text_content
        # HTML内に実際のスタイルが含まれているか
        assert "border-top: 3px solid black" in text_content
        assert "colspan=\"6\"" in text_content

def test_cells_metadata_integrity(tmp_path, complex_xlsx):
    """cells メタデータの内容が期待通りか検証"""
    extractor = MetadataExtractor(output_dir=tmp_path)
    result = extractor.extract(complex_xlsx)
    cells = result["sheets"]["ComplexLayout"]["cells"]
    
    # 結合セルのメタデータ
    a1_cell = next(c for c in cells if c["coord"] == "A1")
    assert a1_cell["colspan"] == 6
    assert a1_cell["rowspan"] == 2
    
    # 罫線のメタデータ
    assert a1_cell["style"]["borders"]["top"] == "thick"
    
    # 値の確認
    b4_cell = next(c for c in cells if c["coord"] == "B4")
    assert b4_cell["value"] == "2026-04-18"
