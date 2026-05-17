import threading
import time
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from kami_excel_extractor.converter import ExcelConverter


def test_concurrent_conversion_isolation(tmp_path):
    """
    複数スレッドから同時に convert を呼び出した際、
    中間ファイルのパスが衝突しないことを検証する。
    """
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    converter = ExcelConverter(output_dir)

    # 共通のファイル名を持つ入力ファイルを用意
    input_file = tmp_path / "report.xlsx"
    input_file.touch()

    active_outdirs = set()
    lock = threading.Lock()
    collision_detected = False

    def mock_run(args, **kwargs):
        nonlocal collision_detected
        if "--convert-to" in args and "pdf" in args:
            outdir = args[args.index("--outdir") + 1]

            with lock:
                if outdir in active_outdirs:
                    collision_detected = True
                active_outdirs.add(outdir)

            # 中間ファイルの生成をシミュレート
            input_stem = Path(args[-1]).stem
            pdf_path = Path(outdir) / f"{input_stem}.pdf"
            pdf_path.touch()

            # 他のスレッドが割り込む隙を作るためのディレイ
            time.sleep(0.1)

            with lock:
                active_outdirs.remove(outdir)

        elif any("pdftocairo" in arg for arg in args) or (len(args) > 1 and args[1] == "-png"):
            # pdftocairo -png prefix ... -> prefix.png を生成する
            prefix = args[-1]
            png_path = Path(f"{prefix}.png")
            png_path.touch()
            print(f"DEBUG: touched {png_path} (args: {args})")

        mock_res = MagicMock()
        mock_res.returncode = 0
        return mock_res

    with patch("subprocess.run", side_effect=mock_run):
        with patch("shutil.which", return_value="/usr/bin/mock"):
            # 2つのスレッドで同時に実行
            threads = []
            for _ in range(2):
                t = threading.Thread(target=converter.convert, args=(input_file,))
                threads.append(t)
                t.start()

            for t in threads:
                t.join()

    assert not collision_detected, "Race condition detected: same outdir used for concurrent conversions"
