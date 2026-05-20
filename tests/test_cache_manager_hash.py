import hashlib
from pathlib import Path

import pytest

from kami_excel_extractor.utils import CacheManager


@pytest.fixture
def cache_manager(tmp_path):
    db_path = tmp_path / "test_cache.db"
    return CacheManager(db_path)


def test_get_file_hash_normal(cache_manager, tmp_path):
    """正常なファイルの内容から正しくハッシュが生成されることを検証"""
    file_path = tmp_path / "normal.txt"
    content = b"Hello, world!"
    file_path.write_bytes(content)

    expected = hashlib.sha256(content).hexdigest()
    assert cache_manager.get_file_hash(file_path) == expected


def test_get_file_hash_empty(cache_manager, tmp_path):
    """空のファイルから正しくハッシュが生成されることを検証"""
    file_path = tmp_path / "empty.txt"
    file_path.write_bytes(b"")

    expected = hashlib.sha256(b"").hexdigest()
    assert cache_manager.get_file_hash(file_path) == expected


def test_get_file_hash_large(cache_manager, tmp_path):
    """バッファサイズ(8192)を超えるファイルから正しくハッシュが生成されることを検証"""
    file_path = tmp_path / "large.bin"
    # 8192 * 2 + 100 bytes
    content = b"a" * (8192 * 2 + 100)
    file_path.write_bytes(content)

    expected = hashlib.sha256(content).hexdigest()
    assert cache_manager.get_file_hash(file_path) == expected


def test_get_file_hash_not_found(cache_manager, tmp_path):
    """存在しないファイルの場合にFileNotFoundErrorが発生することを検証"""
    file_path = tmp_path / "non_existent.txt"
    with pytest.raises(FileNotFoundError):
        cache_manager.get_file_hash(file_path)


@pytest.mark.asyncio
async def test_aget_file_hash(cache_manager, tmp_path):
    """非同期版のaget_file_hashが正しく動作することを検証"""
    file_path = tmp_path / "async.txt"
    content = b"Async test content"
    file_path.write_bytes(content)

    expected = hashlib.sha256(content).hexdigest()
    result = await cache_manager.aget_file_hash(file_path)
    assert result == expected


@pytest.mark.asyncio
async def test_aget_file_hash_not_found(cache_manager, tmp_path):
    """非同期版でも存在しないファイルの場合にFileNotFoundErrorが発生することを検証"""
    file_path = tmp_path / "non_existent_async.txt"
    with pytest.raises(FileNotFoundError):
        await cache_manager.aget_file_hash(file_path)
