import asyncio
import hashlib
from pathlib import Path

import pytest

from kami_excel_extractor.utils import CacheManager, clean_kami_text, secure_filename


def test_clean_kami_text():
    """clean_kami_textの挙動を検証"""
    assert clean_kami_text("氏  名") == "氏名"
    assert clean_kami_text("氏   名") == "氏名"
    assert clean_kami_text("氏    名") == "氏    名"  # 4つ以上は対象外
    assert clean_kami_text("あ　い") == "あい"  # 全角スペース
    assert clean_kami_text("  trimmed  ") == "trimmed"
    assert clean_kami_text(123) == 123  # 非文字列
    assert clean_kami_text(None) is None


@pytest.mark.asyncio
async def test_aget_file_hash(tmp_path):
    """aget_file_hashの非同期挙動を検証"""
    db_path = tmp_path / "test.db"
    cache_manager = CacheManager(db_path)
    file_path = tmp_path / "async_test.txt"
    content = b"async content"
    file_path.write_bytes(content)

    expected_hash = hashlib.sha256(content).hexdigest()
    assert await cache_manager.aget_file_hash(file_path) == expected_hash


def test_secure_filename():
    test_cases = [
        ("NormalName", "NormalName"),
        ("Spaces In Name", "Spaces_In_Name"),
        ("../../../tmp/evil", "tmp_evil"),
        ("..", "unnamed"),
        (".", "unnamed"),
        ("", "unnamed"),
        (None, "unnamed"),
        ("!@#$%^&*()", "unnamed"),
        ("Sheet (1)", "Sheet_1"),
        ("シート１", "シート1"),  # Japanese support
        ("My.File.xlsx", "My.File.xlsx"),
        ("___multiple___", "multiple"),
        ("...dots...", "dots"),
    ]

    for input_str, expected_output in test_cases:
        output = secure_filename(input_str)
        print(f"Input: '{input_str}' -> Output: '{output}' (Expected: '{expected_output}')")
        assert output == expected_output or (output == "unnamed" and expected_output == "unnamed")


def test_secure_filename_edge_cases():
    """secure_filenameの追加のエッジケースを検証"""
    # 非常に長いファイル名
    long_name = "a" * 1000
    assert secure_filename(long_name) == long_name

    # すべてサニタイズされて消える場合
    assert secure_filename("!!!") == "unnamed"


if __name__ == "__main__":
    import os
    import sys

    # Add src to sys.path if needed
    sys.path.append(os.path.join(os.getcwd(), "src"))
    test_secure_filename()
