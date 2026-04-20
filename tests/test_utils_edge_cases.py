import pytest
from kami_excel_extractor.utils import secure_filename, clean_kami_text

@pytest.mark.parametrize("input_val", [
    None,
    123,
    45.67,
    True,
    False,
    ["list"],
    {"dict": "val"},
])
def test_clean_kami_text_non_string(input_val):
    """文字列以外の入力に対する挙動 (そのまま返却されること)"""
    assert clean_kami_text(input_val) == input_val

def test_secure_filename_all_unsafe():
    """全てが安全でない文字で構成されている場合"""
    assert secure_filename("!@#$%^&*()") == "unnamed"

def test_secure_filename_long_extension():
    """長い拡張子や複数のドットが含まれる場合"""
    assert secure_filename("my.report.v1.0.pdf") == "my.report.v1.0.pdf"
    assert secure_filename("data...json") == "data.json"

def test_secure_filename_japanese():
    """日本語を含むファイル名のサニタイズ (NFKD 正規化の挙動を検証)"""
    # 報告書 -> 正常に維持される
    assert secure_filename("報告書_2026.pdf") == "報告書_2026.pdf"
    # スペース混じり、濁点(混) -> NFKD で濁点が分離され、[^\w] で _ に置換される現状の挙動
    # '  スペース 混じり  .xlsx  '
    # 'スヘ_ース_混し_り_.xlsx'
    assert secure_filename("  スペース 混じり  .xlsx  ") == "スヘ_ース_混し_り_.xlsx"

def test_secure_filename_directory_traversal():
    """ディレクトリ・トラバーサル攻撃の無効化"""
    assert secure_filename("../../../etc/passwd") == "etc_passwd"
    assert secure_filename("./local/file") == "local_file"

def test_secure_filename_empty():
    """空文字に対する挙動"""
    assert secure_filename("") == "unnamed"
    assert secure_filename(" ") == "unnamed"

def test_secure_filename_reserved_names():
    """ドット単体などの予約語的入力"""
    assert secure_filename(".") == "unnamed"
    assert secure_filename("..") == "unnamed"
