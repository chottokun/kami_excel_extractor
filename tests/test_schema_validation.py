import pytest
from pydantic import ValidationError
from kami_excel_extractor.schema import ExtractionOptions, RagOptions, ExtractionResult

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

def test_extraction_result_validation():
    """ExtractionResult のバリデーションとデフォルト値を確認"""
    # デフォルト値
    res = ExtractionResult()
    assert res.data is None
    assert res.error is None

    # 辞書データの保持
    res_dict = ExtractionResult(data={"key": "value"})
    assert res_dict.data == {"key": "value"}

    # リストデータの保持
    res_list = ExtractionResult(data=[1, 2, 3])
    assert res_list.data == [1, 2, 3]

    # エラー情報の保持
    res_err = ExtractionResult(error="test error")
    assert res_err.error == "test error"

    # 追加フィールド (_raw_json 等) の許容
    res_extra = ExtractionResult(_raw_json='{"test": 1}', other="field")
    assert res_extra.model_extra["_raw_json"] == '{"test": 1}'
    assert res_extra.model_extra["other"] == "field"
