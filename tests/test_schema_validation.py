import pytest
from pydantic import ValidationError
from kami_excel_extractor.schema import (
    ExtractionOptions,
    RagOptions,
    ExtractionResult,
    SheetData,
    FullExtraction
)

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
    """ExtractionResult のバリデーションと extra='allow' を確認"""
    result = ExtractionResult(data={"key": "value"}, extra_field="extra")
    assert result.data == {"key": "value"}
    assert result.extra_field == "extra"

    # data が List の場合
    result_list = ExtractionResult(data=["item1", "item2"])
    assert result_list.data == ["item1", "item2"]

def test_sheet_data_default():
    """SheetData のデフォルト値を確認"""
    sheet = SheetData()
    assert sheet.metadata == {}
    assert sheet.sections == []
    assert sheet.data == []
    assert sheet.errors == []

def test_sheet_data_validation():
    """SheetData のバリデーションと extra='allow' を確認"""
    sheet = SheetData(
        metadata={"author": "test"},
        sections=[{"title": "Section 1"}],
        data={"main": "content"},
        errors=["error1"],
        unknown_field="allowed"
    )
    assert sheet.metadata == {"author": "test"}
    assert sheet.sections == [{"title": "Section 1"}]
    assert sheet.data == {"main": "content"}
    assert sheet.errors == ["error1"]
    assert sheet.unknown_field == "allowed"

    # 不正な型のバリデーション (metadata は Dict を期待)
    with pytest.raises(ValidationError):
        SheetData(metadata="not-a-dict")

def test_full_extraction_validation():
    """FullExtraction のバリデーションを確認"""
    sheet_data = SheetData(data={"k": "v"})
    full = FullExtraction(sheets={"Sheet1": sheet_data})
    assert "Sheet1" in full.sheets
    assert full.sheets["Sheet1"].data == {"k": "v"}

    # dict からの生成
    full_from_dict = FullExtraction(sheets={"Sheet1": {"data": {"k": "v"}}})
    assert isinstance(full_from_dict.sheets["Sheet1"], SheetData)
    assert full_from_dict.sheets["Sheet1"].data == {"k": "v"}

    # extra='allow'
    full_extra = FullExtraction(sheets={}, extra_info="test")
    assert full_extra.extra_info == "test"
