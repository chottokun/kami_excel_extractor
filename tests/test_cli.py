import json
import logging
import os
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock, AsyncMock, mock_open
import pytest
from kami_excel_extractor.cli import main

@pytest.fixture
def mock_extractor():
    with patch("kami_excel_extractor.cli.KamiExcelExtractor") as mock:
        yield mock

def test_main_success(mock_extractor, caplog):
    """正常なファイル処理フローを確認 (デフォルトモード)"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_structured_data = AsyncMock()
    mock_inst.aextract_structured_data.return_value = {"sheets": {"Sheet1": {"foo": "bar"}}}

    # Mock open and json.dump
    m_open = mock_open()

    with patch("sys.argv", ["kami-excel", "test.xlsx"]), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump") as mock_json_dump:

        with caplog.at_level(logging.INFO):
            main()

        # extractor が正しく初期化されたか
        mock_extractor.assert_called_once()

        # aextract_structured_data が正しく呼ばれたか
        mock_inst.aextract_structured_data.assert_called_once_with(
            "test.xlsx",
            model=None,
            system_prompt=None,
            include_visual_summaries=True,
            use_visual_context=True
        )

        # 結果が保存されたか
        # Path("output") / "test_result.json"
        expected_output_path = Path("output") / "test_result.json"
        m_open.assert_any_call(expected_output_path, "w", encoding="utf-8")
        mock_json_dump.assert_called()

        assert "結果を保存しました" in caplog.text

def test_main_rag_success(mock_extractor, caplog):
    """RAGモードでの正常処理を確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_rag_chunks = AsyncMock()

    # aextract_rag_chunks returns (chunks_map, structured_data)
    dummy_chunks_map = {"Sheet1": {"chunks": [{"chunks": []}], "markdown": "# Content"}}
    dummy_structured_data = {"sheets": {"Sheet1": {}}}
    mock_inst.aextract_rag_chunks.return_value = (dummy_chunks_map, dummy_structured_data)

    m_open = mock_open()

    with patch("sys.argv", ["kami-excel", "test.xlsx", "--rag"]), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump") as mock_json_dump:

        with caplog.at_level(logging.INFO):
            main()

        # aextract_rag_chunks が呼ばれたか
        mock_inst.aextract_rag_chunks.assert_called_once_with(
            "test.xlsx",
            model=None,
            use_visual_context=True
        )

        # 結果とRAG用データが保存されたか
        expected_result_path = Path("output") / "test_result.json"
        expected_rag_path = Path("output") / "test_rag.json"

        m_open.assert_any_call(expected_result_path, "w", encoding="utf-8")
        m_open.assert_any_call(expected_rag_path, "w", encoding="utf-8")

        assert "RAG用データを保存しました" in caplog.text

def test_main_file_not_found(caplog):
    """入力ファイルが見つからない場合のエラー終了を確認"""
    with patch("sys.argv", ["kami-excel", "non_existent.xlsx"]), \
         patch("pathlib.Path.exists", return_value=False), \
         patch("asyncio.run"), \
         patch("sys.exit", side_effect=SystemExit(1)) as mock_exit:

        with caplog.at_level(logging.ERROR):
            with pytest.raises(SystemExit):
                main()

        mock_exit.assert_called_once_with(1)
        assert "入力ファイルが見つかりません" in caplog.text

def test_main_extraction_error(mock_extractor, caplog):
    """抽出中に例外が発生した場合のエラー終了を確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_structured_data = AsyncMock()
    mock_inst.aextract_structured_data.side_effect = Exception("Extraction failed")

    with patch("sys.argv", ["kami-excel", "test.xlsx"]), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("sys.exit", side_effect=SystemExit(1)) as mock_exit:

        with caplog.at_level(logging.ERROR):
            with pytest.raises(SystemExit):
                main()

        mock_exit.assert_called_once_with(1)
        assert "実行中にエラーが発生しました: Extraction failed" in caplog.text

def test_main_args_passing(mock_extractor):
    """引数が正しく渡されることを確認"""
    mock_inst = mock_extractor.return_value
    mock_inst.aextract_structured_data = AsyncMock()
    mock_inst.aextract_structured_data.return_value = {}

    test_args = [
        "kami-excel", "test.xlsx",
        "--model", "gpt-4",
        "--timeout", "120",
        "--rpm", "10",
        "--no-vision",
        "--api-key", "secret-key",
        "--base-url", "https://api.openai.com",
        "--output-dir", "custom_out"
    ]

    m_open = mock_open()

    with patch("sys.argv", test_args), \
         patch("pathlib.Path.exists", return_value=True), \
         patch("builtins.open", m_open), \
         patch("json.dump"), \
         patch.dict(os.environ, {}, clear=False):

        main()

        # extractor の初期化引数を確認
        mock_extractor.assert_called_once_with(
            api_key="secret-key",
            base_url="https://api.openai.com",
            output_dir="custom_out",
            timeout=120.0
        )

        # RPMが環境変数にセットされたか
        assert os.environ.get("LLM_RPM_LIMIT") == "10"

        # aextract_structured_data の引数を確認 (--no-vision の影響)
        mock_inst.aextract_structured_data.assert_called_once_with(
            "test.xlsx",
            model="gpt-4",
            system_prompt=None,
            include_visual_summaries=False,
            use_visual_context=False
        )
