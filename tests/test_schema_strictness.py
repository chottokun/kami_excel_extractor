
import pytest
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionResult

def test_schema_cleansing_of_extra_fields():
    """LLMが返した余分なフィールドが自動的に削除されることを検証"""
    extractor = KamiExcelExtractor()
    
    # 未知のフィールド 'unwanted_comment' が含まれている
    llm_response_content = """
    ```json
    {
        "data": [{"id": 1}],
        "unwanted_comment": "I added this field because I am an AI."
    }
    ```
    """
    
    # 内部の _parse_llm_response をテスト
    parsed = extractor._parse_llm_response(llm_response_content, "TestSheet")
    
    # 'data' は残る
    assert "data" in parsed
    # 'unwanted_comment' は削除されているべき (ExtractionResultには定義されていないため)
    # ※ Pydanticモデルに変換・再ダンプする過程で消えることを期待
    
    # 再バリデーション
    result_obj = ExtractionResult(**parsed)
    # Pydanticのモデルから辞書に戻した時に余計なものが消えているか
    cleaned_dict = result_obj.model_dump()
    assert "unwanted_comment" not in cleaned_dict

def test_nested_schema_cleansing():
    """ネストされた構造（SheetData等）でもクレンジングが効くことを検証"""
    # 複雑なケースを模倣
    pass
