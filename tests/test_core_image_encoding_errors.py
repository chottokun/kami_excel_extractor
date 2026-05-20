import asyncio
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

from kami_excel_extractor.core import KamiExcelExtractor


@pytest.fixture
def extractor(tmp_path):
    return KamiExcelExtractor(output_dir=tmp_path)


@pytest.mark.asyncio
async def test_encode_image_file_not_found(extractor):
    """ファイルが存在しない場合の挙動を確認"""
    invalid_path = Path("non_existent_image.png")
    with pytest.raises(FileNotFoundError):
        await extractor._encode_image_to_base64_url(invalid_path)


@pytest.mark.asyncio
async def test_encode_image_permission_error(extractor, tmp_path):
    """ファイルが読み取れない場合の挙動を確認"""
    image_path = tmp_path / "noperm.png"
    image_path.write_text("dummy")

    with patch("builtins.open", side_effect=PermissionError("Permission denied")):
        # Note: aget_file_hash uses built-in open via to_thread
        with pytest.raises(PermissionError):
            await extractor._encode_image_to_base64_url(image_path)


@pytest.mark.asyncio
async def test_aget_visual_summary_encoding_failure(extractor, tmp_path):
    """画像エンコード失敗時の aget_visual_summary の挙動を確認"""
    image_path = tmp_path / "test.png"
    image_path.write_text("dummy")

    # _encode_image_to_base64_url が失敗するようにモック
    with patch.object(
        KamiExcelExtractor, "_encode_image_to_base64_url", side_effect=FileNotFoundError("File not found")
    ):
        summary = await extractor.aget_visual_summary(image_path)
        # 修正後は例外をキャッチして fallback メッセージを返すはず
        assert summary == "[画像概要] 解析失敗。"


@pytest.mark.asyncio
async def test_aprocess_chart_data_encoding_failure(extractor, tmp_path):
    """画像エンコード失敗時の _aprocess_chart_data の挙動を確認"""
    media_item = {"filename": "chart.png"}
    (extractor.output_dir / "media").mkdir(parents=True, exist_ok=True)
    image_path = extractor.output_dir / "media" / "chart.png"
    image_path.write_text("dummy")

    with patch.object(KamiExcelExtractor, "_encode_image_to_base64_url", side_effect=Exception("Encoding error")):
        # 修正後は例外をキャッチして media_item をそのまま返すはず
        result = await extractor._aprocess_chart_data(media_item, "model", None)
        assert result == media_item
        assert "visual_data" not in result
