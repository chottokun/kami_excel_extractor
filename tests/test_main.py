import pytest
import logging
from unittest.mock import patch, MagicMock
from pathlib import Path
import main

def test_main_no_api_key(caplog):
    """GEMINI_API_KEYが設定されていない場合にエラーログを出力して終了することを確認"""
    with patch("main.LLM_API_KEY", None), \
         patch("main.LLM_MODEL", "gemini/gemini-1.5-flash"), \
         patch("main.INPUT_DIR", Path("/tmp/empty_dir")), \
         patch("os.getenv", return_value=None):
        with caplog.at_level(logging.ERROR):
            main.main()
        assert "LLM_API_KEY or GEMINI_API_KEY is not set." in caplog.text

def test_main_success(tmp_path, caplog):
    """正常なファイル処理フローを確認"""
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    input_dir.mkdir()
    output_dir.mkdir()

    # テスト用のExcelファイルを作成
    test_file = input_dir / "test.xlsx"
    test_file.touch()

    sheet_results = {
        "Sheet1": {
            "structured": {"foo": "bar"},
            "yaml": "foo: bar",
            "chunks": [{"chunk": 1}],
            "markdown": "# Sheet1 Content"
        }
    }
    full_structured_data = {"all": "data"}

    with patch("main.LLM_API_KEY", "fake_key"), \
         patch("main.INPUT_DIR", input_dir), \
         patch("main.OUTPUT_DIR", output_dir), \
         patch("main.KamiExcelExtractor") as mock_extractor_cls, \
         patch("main.time.sleep", side_effect=KeyboardInterrupt):

        mock_extractor = mock_extractor_cls.return_value
        mock_extractor.extract_rag_chunks.return_value = (sheet_results, full_structured_data)
        mock_extractor.doc_generator = MagicMock()

        with caplog.at_level(logging.INFO):
            # 無限ループをKeyboardInterruptで抜ける
            with pytest.raises(KeyboardInterrupt):
                main.main()

        # 出力ディレクトリとファイルが作成されたか
        target_dir = output_dir / "test"
        assert target_dir.exists()
        assert (target_dir / "full_lib_result.json").exists()
        assert (target_dir / "Sheet1_lib_result.json").exists()
        assert (target_dir / "Sheet1_rag_chunks.json").exists()

        # 成功ログが出力されたか
        assert f"Success: Outputs saved to {target_dir}" in caplog.text

def test_main_exception(tmp_path, caplog):
    """処理中に例外が発生した場合にエラーログを出力し、ループを継続することを確認"""
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    input_dir.mkdir()
    output_dir.mkdir()

    test_file = input_dir / "test.xlsx"
    test_file.touch()

    with patch("main.LLM_API_KEY", "fake_key"), \
         patch("main.INPUT_DIR", input_dir), \
         patch("main.OUTPUT_DIR", output_dir), \
         patch("main.KamiExcelExtractor") as mock_extractor_cls, \
         patch("main.time.sleep", side_effect=KeyboardInterrupt):

        mock_extractor = mock_extractor_cls.return_value
        mock_extractor.extract_rag_chunks.side_effect = Exception("Test Error")

        with caplog.at_level(logging.ERROR):
            with pytest.raises(KeyboardInterrupt):
                main.main()

        # caplog経由でのエラーログ確認
        assert "Failed to process test.xlsx: Test Error" in caplog.text
