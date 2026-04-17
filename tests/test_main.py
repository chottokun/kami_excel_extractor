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

def test_main_partial_failure(tmp_path, caplog):
    """一部のファイルが失敗しても、他のファイルが処理されることを確認"""
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "output"
    input_dir.mkdir()
    output_dir.mkdir()

    # 2つのファイルを用意
    f1 = input_dir / "fail.xlsx"
    f1.touch()
    f2 = input_dir / "success.xlsx"
    f2.touch()

    # globの結果を固定するためにパッチを当てる
    # 順番を保証するためにリストにする
    with patch("main.LLM_API_KEY", "fake_key"), \
         patch("main.INPUT_DIR") as mock_input_dir, \
         patch("main.OUTPUT_DIR", output_dir), \
         patch("main.KamiExcelExtractor") as mock_extractor_cls, \
         patch("main.time.sleep", side_effect=KeyboardInterrupt):

        mock_input_dir.glob.return_value = [f1, f2]
        mock_extractor = mock_extractor_cls.return_value

        # 1回目は失敗、2回目は成功をシミュレート
        mock_extractor.extract_rag_chunks.side_effect = [
            Exception("Fail 1"),
            ({"Sheet1": {"structured": {}, "yaml": "", "chunks": [], "markdown": ""}}, {})
        ]
        mock_extractor.doc_generator = MagicMock()

        with caplog.at_level(logging.INFO):
            with pytest.raises(KeyboardInterrupt):
                main.main()

        # 1つ目のファイルのエラーが記録されていること
        assert "Failed to process fail.xlsx: Fail 1" in caplog.text
        # 2つ目のファイルが成功していること
        assert "Success: Outputs saved to" in caplog.text
        assert "success" in caplog.text

def test_main_mkdir_failure(tmp_path, caplog):
    """ディレクトリ作成に失敗した場合にエラーログを出力することを確認"""
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
         patch("main.Path.mkdir", side_effect=OSError("Mock OSError")), \
         patch("main.time.sleep", side_effect=KeyboardInterrupt):

        mock_extractor = mock_extractor_cls.return_value
        mock_extractor.extract_rag_chunks.return_value = ({"Sheet1": {"structured": {}, "yaml": "", "chunks": [], "markdown": ""}}, {})

        with caplog.at_level(logging.ERROR):
            with pytest.raises(KeyboardInterrupt):
                main.main()

        # OSErrorがキャッチされ、エラーログが出力されていること
        assert "Failed to process test.xlsx: Mock OSError" in caplog.text

def test_main_pdf_generation_failure(tmp_path, caplog):
    """PDF生成に失敗した場合にエラーログを出力することを確認"""
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
        mock_extractor.extract_rag_chunks.return_value = ({"Sheet1": {"structured": {}, "yaml": "", "chunks": [], "markdown": ""}}, {})
        # generate_pdfで例外を投げるように設定
        mock_extractor.doc_generator.generate_pdf.side_effect = Exception("PDF Generation Error")

        with caplog.at_level(logging.ERROR):
            with pytest.raises(KeyboardInterrupt):
                main.main()

        # 例外がキャッチされ、エラーログが出力されていること
        assert "Failed to process test.xlsx: PDF Generation Error" in caplog.text
