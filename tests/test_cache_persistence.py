
import pytest
import sqlite3
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor
from unittest.mock import patch, AsyncMock

@pytest.fixture
def cache_db_path(tmp_path):
    return tmp_path / "test_cache.db"

@pytest.mark.asyncio
async def test_cache_persistence_across_sessions(tmp_path, cache_db_path):
    """セッションを跨いでキャッシュが永続化されることを検証"""
    img_path = tmp_path / "persistent_test.png"
    img_path.write_bytes(b"consistent content")
    
    # セッション1: 初回解析
    extractor1 = KamiExcelExtractor(output_dir=tmp_path)
    # 内部的に使うDBパスを固定
    extractor1._db_path = cache_db_path
    
    with patch("litellm.acompletion") as mock_completion:
        mock_completion.return_value.choices = [
            AsyncMock(message=AsyncMock(content="Analysis Result v1"))
        ]
        res1 = await extractor1.aget_visual_summary(img_path)
        assert res1 == "Analysis Result v1"
        assert mock_completion.call_count == 1

    # セッション2: 新しいインスタンスで再実行
    extractor2 = KamiExcelExtractor(output_dir=tmp_path)
    extractor2._db_path = cache_db_path
    
    with patch("litellm.acompletion") as mock_completion2:
        # モックは呼ばれないはず（キャッシュから返るため）
        res2 = await extractor2.aget_visual_summary(img_path)
        assert res2 == "Analysis Result v1"
        assert mock_completion2.call_count == 0

@pytest.mark.asyncio
async def test_cache_invalidation_on_content_change(tmp_path, cache_db_path):
    """画像の中身が変わった場合にキャッシュが効かない（再解析される）ことを検証"""
    img_path = tmp_path / "changing.png"
    
    extractor = KamiExcelExtractor(output_dir=tmp_path)
    extractor._db_path = cache_db_path
    
    # 1回目
    img_path.write_bytes(b"content A")
    with patch("litellm.acompletion") as mock_completion:
        mock_completion.return_value.choices = [AsyncMock(message=AsyncMock(content="Result A"))]
        await extractor.aget_visual_summary(img_path)
    
    # 2回目: 中身を変更
    img_path.write_bytes(b"content B")
    with patch("litellm.acompletion") as mock_completion:
        mock_completion.return_value.choices = [AsyncMock(message=AsyncMock(content="Result B"))]
        res = await extractor.aget_visual_summary(img_path)
        assert res == "Result B"
        assert mock_completion.call_count == 1
