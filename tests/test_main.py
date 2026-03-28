import pytest
import logging
from pathlib import Path
from unittest.mock import patch, MagicMock
from src.main import main

def test_main_processing_error(caplog):
    """メインループ内での処理失敗時のエラーハンドリングをテストする"""

    # GEMINI_API_KEYをセットして早期リターンを回避
    with patch("src.main.GEMINI_API_KEY", "fake_key"):
        # INPUT_DIR をモック化
        with patch("src.main.INPUT_DIR") as mock_input_dir:
            # glob を呼び出した時に、テスト用のファイルを1つ返すようにする
            mock_input_dir.glob.return_value = [Path("error.xlsx")]

            # KamiExcelExtractor をモック化
            with patch("src.main.KamiExcelExtractor") as mock_extractor_class:
                mock_extractor = mock_extractor_class.return_value
                # extract_rag_chunks が例外を投げるように設定
                mock_extractor.extract_rag_chunks.side_effect = Exception("test error")

                # time.sleep をモック化して KeyboardInterrupt を投げ、無限ループを抜けるようにする
                with patch("src.main.time.sleep", side_effect=KeyboardInterrupt):
                    # caplogでINFO以上のログをキャプチャ
                    with caplog.at_level(logging.INFO):
                        try:
                            main()
                        except KeyboardInterrupt:
                            # ループを抜けるための想定内の例外
                            pass

    # 期待されるエラーログが出力されているか確認
    assert "Failed to process error.xlsx: test error" in caplog.text
    assert "Processing: error.xlsx" in caplog.text
