import sqlite3
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from kami_excel_extractor.utils import CacheManager

def test_init_db_mkdir_failure(tmp_path):
    """Path.mkdir が失敗した場合のテスト"""
    db_path = tmp_path / "subdir" / "cache.db"

    # Note: We need to be careful with patching Path.mkdir as it might be used by pytest/tmp_path
    # but CacheManager calls it on self.db_path.parent
    with patch.object(Path, "mkdir", side_effect=OSError("Permission denied")):
        with pytest.raises(OSError, match="Permission denied"):
            CacheManager(db_path)

def test_init_db_connect_failure(tmp_path):
    """sqlite3.connect が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    with patch("sqlite3.connect", side_effect=sqlite3.Error("Connection failed")):
        with pytest.raises(sqlite3.Error, match="Connection failed"):
            CacheManager(db_path)

def test_init_db_pragma_failure(tmp_path):
    """PRAGMA の設定が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    # Mock connection and its execute method
    mock_conn = MagicMock()
    mock_conn.execute.side_effect = sqlite3.Error("Pragma failed")

    with patch("sqlite3.connect", return_value=mock_conn):
        with pytest.raises(sqlite3.Error, match="Pragma failed"):
            CacheManager(db_path)

def test_init_db_create_table_failure(tmp_path):
    """テーブル作成が失敗した場合のテスト"""
    db_path = tmp_path / "cache.db"

    mock_conn = MagicMock()
    # The first call to execute is PRAGMA, subsequent ones are CREATE TABLE
    # We want to fail on CREATE TABLE
    def side_effect(sql, *args):
        if "CREATE TABLE" in sql:
            raise sqlite3.Error("Create table failed")
        return MagicMock()

    mock_conn.execute.side_effect = side_effect

    with patch("sqlite3.connect", return_value=mock_conn):
        with pytest.raises(sqlite3.Error, match="Create table failed"):
            CacheManager(db_path)
