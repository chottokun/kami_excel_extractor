import subprocess
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from kami_excel_extractor.converter import ExcelConverter


@pytest.fixture
def converter(tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    return ExcelConverter(output_dir)


@pytest.fixture
def pdf_path(tmp_path):
    path = tmp_path / "test.pdf"
    path.touch()
    return path


@pytest.fixture
def mock_run():
    with patch("subprocess.run") as mock:
        yield mock


@pytest.fixture
def mock_which():
    with patch("shutil.which") as mock:
        yield mock


def test_convert_pdf_to_multi_png_pdftocairo_subprocess_error(converter, pdf_path, tmp_path, mock_which, mock_run):
    """pdftocairo が存在するが、実行中に失敗した場合に fitz にフォールバックするテスト"""
    output_prefix = tmp_path / "test_multi"
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "pdftocairo" else None

    # pdftocairo の失敗をシミュレート (check=True は CalledProcessError を発生させる)
    mock_run.side_effect = subprocess.CalledProcessError(1, "pdftocairo")

    with patch.object(converter, "_try_fitz_multi") as mock_fitz_multi:
        mock_fitz_multi.return_value = [Path("fake_fitz.png")]
        results = converter._convert_pdf_to_multi_png(pdf_path, output_prefix)

        assert results == [Path("fake_fitz.png")]
        mock_fitz_multi.assert_called_once_with(pdf_path, output_prefix)


def test_convert_pdf_to_multi_png_pdftocairo_no_files_generated(converter, pdf_path, tmp_path, mock_which, mock_run):
    """pdftocairo は成功したが、PNG ファイルが生成されなかった場合に fitz にフォールバックするテスト"""
    output_prefix = tmp_path / "test_multi"
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "pdftocairo" else None

    # 正常終了したがファイル生成されなかった状況をシミュレート
    mock_run.return_value = MagicMock(returncode=0)

    with patch.object(converter, "_try_fitz_multi") as mock_fitz_multi:
        mock_fitz_multi.return_value = [Path("fake_fitz_no_files.png")]
        results = converter._convert_pdf_to_multi_png(pdf_path, output_prefix)

        assert results == [Path("fake_fitz_no_files.png")]
        mock_fitz_multi.assert_called_once_with(pdf_path, output_prefix)


def test_convert_pdf_to_multi_png_pdftocairo_general_exception(converter, pdf_path, tmp_path, mock_which, mock_run):
    """pdftocairo の実行中に想定外の例外が発生した場合に fitz にフォールバックするテスト"""
    output_prefix = tmp_path / "test_multi"
    mock_which.side_effect = lambda x: f"/usr/bin/{x}" if x == "pdftocairo" else None

    # 想定外の OSError をシミュレート
    mock_run.side_effect = OSError("Unexpected failure")

    with patch.object(converter, "_try_fitz_multi") as mock_fitz_multi:
        mock_fitz_multi.return_value = [Path("fake_fitz_exception.png")]
        results = converter._convert_pdf_to_multi_png(pdf_path, output_prefix)

        assert results == [Path("fake_fitz_exception.png")]
        mock_fitz_multi.assert_called_once_with(pdf_path, output_prefix)


def test_convert_file_size_limit_exceeded(converter, tmp_path):
    """ファイルサイズが上限を超えている場合に ValueError が発生するテスト"""
    input_file = tmp_path / "large.xlsx"
    input_file.touch()

    # ファイルサイズを 51MB (上限 50MB) にシミュレート
    with patch("pathlib.Path.stat") as mock_stat:
        mock_stat.return_value.st_size = 51 * 1024 * 1024
        with pytest.raises(ValueError, match="exceeds the limit"):
            converter.convert(input_file)


def test_convert_soffice_not_found(converter, tmp_path, mock_which):
    """soffice が PATH に見つからない場合に RuntimeError が発生するテスト"""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    mock_which.side_effect = lambda x: None

    with pytest.raises(RuntimeError, match=r"LibreOffice \(soffice\) not found in PATH"):
        converter.convert(input_file)


def test_convert_with_sheet_name(converter, tmp_path, mock_which, mock_run):
    """シート指定のある変換テスト（対象シートのみを孤立化する処理の検証を含む）"""
    input_file = tmp_path / "test.xlsx"
    input_file.touch()

    mock_which.side_effect = lambda x: f"/usr/bin/{x}"

    # openpyxl をモック化して実際ファイルの書き込み等をスキップ
    with patch("openpyxl.load_workbook") as mock_load:
        mock_wb = MagicMock()
        mock_wb.sheetnames = ["Sheet1", "Sheet2"]
        mock_load.return_value = mock_wb

        # LibreOfficeの実行完了を模倣して PDF ファイルを生成
        def mock_run_impl(args, **kwargs):
            args_list = list(args)
            if any("soffice" in str(a) for a in args_list):
                try:
                    outdir_idx = args_list.index("--outdir")
                    outdir = Path(args_list[outdir_idx + 1])
                    input_path = Path(args_list[-1])
                    (outdir / f"{input_path.stem}.pdf").touch()
                except (ValueError, IndexError):
                    pass
            return MagicMock(returncode=0)

        mock_run.side_effect = mock_run_impl

        with patch.object(converter, "_convert_pdf_to_multi_png") as mock_multi:
            mock_multi.return_value = [Path("sheet1-1.png")]

            result = converter.convert(input_file, sheet_name="Sheet1")

            assert result == [Path("sheet1-1.png")]
            mock_multi.assert_called_once()
            mock_load.assert_called_once_with(input_file, data_only=True)

            # レビュー指摘：非ターゲットシート（Sheet2）が wb から del されているか確認
            # del wb[name] は wb.__delitem__(name) を呼び出すため、これをアサートする
            mock_wb.__delitem__.assert_any_call("Sheet2")
            # 対象シート（Sheet1）は削除されていないことを確認
            with pytest.raises(AssertionError):
                mock_wb.__delitem__.assert_any_call("Sheet1")
