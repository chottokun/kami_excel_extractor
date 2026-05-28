import sqlite3
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from kami_excel_extractor.utils import CacheManager


def test_init_db_mkdir_failure():
    """Path.mkdir が失敗した場合のテスト（グローバルパッチを避け、モックオブジェクトの注入で安全にテスト）"""
    # Pathオブジェクトをモック化
    mock_db_path = MagicMock(spec=Path)
    mock_parent = MagicMock(spec=Path)
    # mkdir呼び出し時にPermissionError（OSError）をシミュレート
    mock_parent.mkdir.side_effect = OSError("Permission denied")
    mock_db_path.parent = mock_parent

    with pytest.raises(OSError, match="Permission denied"):
        CacheManager(mock_db_path)


def test_init_db_connect_failure(tmp_path):
    """sqlite3.connect が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    # connectの呼び出し自体をpatchするが、utils内の対象モジュールのみにスコープ限定
    with patch("kami_excel_extractor.utils.sqlite3.connect", side_effect=sqlite3.Error("Connection failed")):
        with pytest.raises(sqlite3.Error, match="Connection failed"):
            CacheManager(db_path)


def test_init_db_pragma_failure(tmp_path):
    """PRAGMA の設定が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    mock_conn = MagicMock()
    mock_conn.execute.side_effect = sqlite3.Error("Pragma failed")

    with patch("kami_excel_extractor.utils.sqlite3.connect", return_value=mock_conn):
        with pytest.raises(sqlite3.Error, match="Pragma failed"):
            CacheManager(db_path)


def test_init_db_create_table_failure(tmp_path):
    """テーブル作成が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    mock_conn = MagicMock()

    # 1回目のexecuteはPRAGMA用なのでスルーし、CREATE TABLEの時にエラーを投げる
    def side_effect(sql, *args):
        if "CREATE TABLE" in sql:
            raise sqlite3.Error("Create table failed")
        return MagicMock()

    mock_conn.execute.side_effect = side_effect

    with patch("kami_excel_extractor.utils.sqlite3.connect", return_value=mock_conn):
        with pytest.raises(sqlite3.Error, match="Create table failed"):
            CacheManager(db_path)
