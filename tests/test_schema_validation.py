import pytest
from pydantic import ValidationError
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

def test_extraction_options_default():
    """デフォルト値が正しく設定されることを確認"""
    options = ExtractionOptions()
    assert options.model is None
    # 現在の実装では False
    assert options.include_visual_summaries is False
    assert options.use_visual_context is True
    assert options.system_prompt is None

def test_extraction_options_validation():
    """不正な型を与えた際にバリデーションエラーが出ることを確認"""
    with pytest.raises(ValidationError):
        ExtractionOptions(include_visual_summaries="not-a-bool")

def test_rag_options_default():
    """RagOptions のデフォルト値を確認"""
    options = RagOptions()
    assert options.model is None
    assert options.use_visual_context is True
    # 現在の実装では "kv"
    assert options.list_format == "kv"
