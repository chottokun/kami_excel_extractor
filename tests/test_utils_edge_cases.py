import pytest

from kami_excel_extractor.utils import clean_kami_text, secure_filename


class CustomObj:
    pass


@pytest.mark.parametrize(
    "input_val",
    [
        None,
        123,
        45.67,
        True,
        False,
        ["list"],
        {"dict": "val"},
        (1, 2),
        {1, 2},
        b"bytes",
        CustomObj(),
    ],
)
def test_clean_kami_text_non_string(input_val):
    """文字列以外の入力に対する挙動 (そのまま返却されること)"""
    assert clean_kami_text(input_val) == input_val


@pytest.mark.parametrize(
    "input_val, expected",
    [
        ("", ""),
        ("   ", ""),
        ("\n\t", ""),
        ("A" * 1000, "A" * 1000),
    ],
)
def test_clean_kami_text_string_edge_cases(input_val, expected):
    """文字列の境界条件に対する挙動"""
    assert clean_kami_text(input_val) == expected


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


def test_secure_filename_extreme_length():
    """非常に長いファイル名の処理 (10,000文字)"""
    long_name = "a" * 10000
    assert secure_filename(long_name) == long_name


def test_secure_filename_long_unsafe_collapsing():
    """長い不安全な文字列のサニタイズと集約 (10,000文字)"""
    # 記号の連続が適切に処理され、空になった場合は "unnamed" になること
    long_unsafe = "!" * 10000
    assert secure_filename(long_unsafe) == "unnamed"

    # 有効な文字に挟まれた長い記号列が1つのアンダースコアに集約されること
    long_mixed = "a" + "!" * 10000 + "b"
    assert secure_filename(long_mixed) == "a_b"

    # ドットの連続が集約されること
    long_dots = "a" + "." * 10000 + "b"
    assert secure_filename(long_dots) == "a.b"


@pytest.mark.parametrize(
    "input_text, expected",
    [
        # 基本的なケース (1-3個の空白削除)
        ("氏 名", "氏名"),
        ("氏  名", "氏名"),
        ("氏   名", "氏名"),
        # 4個以上の空白 (仕様上削除されない)
        ("氏    名", "氏    名"),
        ("氏     名", "氏     名"),
        # 全角スペース
        ("氏　名", "氏名"),
        ("氏　　名", "氏名"),
        # 混合
        ("氏 　名", "氏名"),
        # 漢字・ひらがな・カタカナの組み合わせ
        ("氏 めい", "氏めい"),
        ("あ　イ", "あイ"),
        ("カ タ 漢", "カタ漢"),
        # 非CJK文字との境界 (削除されない)
        ("氏 Name", "氏 Name"),
        ("Name 名", "Name 名"),
        ("123 名", "123 名"),
        # 文頭・文末の空白 (stripされる)
        ("  氏名  ", "氏名"),
        ("\t氏名\n", "氏名"),
        # 特殊な空白 (タブや改行) - \s に含まれる
        ("氏\t名", "氏名"),
        ("氏\n名", "氏名"),
        # 空文字・空白のみ
        ("", ""),
        ("   ", ""),
        ("　　　", ""),
    ],
)
def test_clean_kami_text_edge_cases(input_text, expected):
    """clean_kami_text の様々なエッジケースを検証"""
    assert clean_kami_text(input_text) == expected
