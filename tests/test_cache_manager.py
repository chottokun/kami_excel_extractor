import pytest
import sqlite3
from pathlib import Path
from kami_excel_extractor.utils import CacheManager

@pytest.fixture
def cache_manager(tmp_path):
    db_path = tmp_path / "test_cache.db"
    return CacheManager(db_path)

def test_llm_result_roundtrip(cache_manager):
    """set_llm_resultで保存した内容がget_llm_resultで取得できることを検証"""
    model = "gpt-4o"
    prompt = "Translate to English"
    input_text = "こんにちは"
    content = "Hello"

    cache_manager.set_llm_result(model, prompt, input_text, content)
    result = cache_manager.get_llm_result(model, prompt, input_text)

    assert result == content

def test_llm_result_not_found(cache_manager):
    """キャッシュに存在しない場合にNoneが返ることを検証"""
    result = cache_manager.get_llm_result("model", "prompt", "input")
    assert result is None

def test_llm_result_overwrite(cache_manager):
    """同じキーで保存した場合に上書きされることを検証"""
    model = "gpt-4o"
    prompt = "Translate"
    input_text = "A"

    cache_manager.set_llm_result(model, prompt, input_text, "First")
    cache_manager.set_llm_result(model, prompt, input_text, "Second")

    result = cache_manager.get_llm_result(model, prompt, input_text)
    assert result == "Second"

def test_llm_result_unique_keys(cache_manager):
    """model, prompt, input_textのいずれかが異なれば別々のキャッシュとして扱われることを検証"""
    model = "m"
    prompt = "p"
    input_text = "i"
    content = "c"

    cache_manager.set_llm_result(model, prompt, input_text, content)

    # modelが違う
    assert cache_manager.get_llm_result("other", prompt, input_text) is None
    # promptが違う
    assert cache_manager.get_llm_result(model, "other", input_text) is None
    # input_textが違う
    assert cache_manager.get_llm_result(model, prompt, "other") is None

def test_vlm_result_roundtrip(cache_manager):
    """set_vlm_resultとget_vlm_resultの挙動を検証"""
    model = "gpt-4o"
    prompt = "Describe image"
    img_hash = "hash123"
    content = "A nice cat"

    cache_manager.set_vlm_result(model, prompt, img_hash, content)
    assert cache_manager.get_vlm_result(model, prompt, img_hash) == content
    assert cache_manager.get_vlm_result(model, prompt, "other") is None

def test_image_data_url_roundtrip(cache_manager):
    """set_image_data_urlとget_image_data_urlの挙動を検証"""
    img_hash = "hash456"
    data_url = "data:image/png;base64,..."

    cache_manager.set_image_data_url(img_hash, data_url)
    assert cache_manager.get_image_data_url(img_hash) == data_url
    assert cache_manager.get_image_data_url("other") is None

def test_clear_cache(cache_manager):
    """clear()がすべてのテーブルをクリアすることを検証"""
    cache_manager.set_llm_result("m", "p", "i", "c")
    cache_manager.set_vlm_result("m", "p", "h", "c")
    cache_manager.set_image_data_url("h", "d")

    cache_manager.clear()

    assert cache_manager.get_llm_result("m", "p", "i") is None
    assert cache_manager.get_vlm_result("m", "p", "h") is None
    assert cache_manager.get_image_data_url("h") is None

def test_get_file_hash(cache_manager, tmp_path):
    """get_file_hashが正しいSHA-256ハッシュを返すことを検証"""
    file_path = tmp_path / "test.txt"
    content = b"test content"
    file_path.write_bytes(content)

    import hashlib
    expected_hash = hashlib.sha256(content).hexdigest()

    assert cache_manager.get_file_hash(file_path) == expected_hash
