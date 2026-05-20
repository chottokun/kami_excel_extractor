import logging
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path
from typing import List, Optional, Union

import openpyxl

from .utils import secure_filename

logger = logging.getLogger(__name__)


class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""

    def __init__(self, output_dir: Path, dpi: int = 150, max_file_size_mb: int = 50):
        self.output_dir = Path(output_dir).resolve()
        self.dpi = dpi
        self.max_file_size_mb = max_file_size_mb

    def convert(self, input_file: Path, sheet_name: Optional[str] = None) -> Union[Path, List[Path]]:
        input_file = input_file.resolve()

        # 🔒 Race Condition Fix: UUIDを使用して中間ファイル名の衝突を回避
        run_id = uuid.uuid4().hex[:12]
        safe_stem = secure_filename(input_file.stem)
        output_prefix = self.output_dir / f"{safe_stem}_{run_id}"

        # 入力ファイルの存在確認
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        # 🔒 Security Fix: ファイルサイズ制限のチェック
        stat_result = input_file.stat()
        file_size_mb = stat_result.st_size / (1024 * 1024)
        if file_size_mb > self.max_file_size_mb:
            raise ValueError(f"File size ({file_size_mb:.1f}MB) exceeds the limit ({self.max_file_size_mb:.1f}MB).")

        with tempfile.TemporaryDirectory(prefix="lo_profile_") as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str).resolve()

            # ターゲットファイルの準備
            # 常に一時ディレクトリにコピーまたは隔離することで、元ファイルへの副作用を完全に排除
            temp_input = tmp_dir / f"input_{run_id}.xlsx"
            if sheet_name:
                logger.info(f"Isolating sheet '{sheet_name}' for conversion...")
                wb = openpyxl.load_workbook(input_file, data_only=True)
                for name in wb.sheetnames:
                    if name != sheet_name:
                        del wb[name]
                wb.save(temp_input)
            else:
                shutil.copy2(input_file, temp_input)

            # LibreOfficeは入力ファイルのステム名でPDFを出力するため、移動先を定義
            original_pdf = tmp_dir / f"converted_{run_id}.pdf"
            expected_pdf = tmp_dir / f"{temp_input.stem}.pdf"

            try:
                # Step 1: Excel -> PDF
                logger.info(f"Converting {input_file.name} to PDF (ID: {run_id})...")

                raw_cmd_path = shutil.which("soffice")
                if not raw_cmd_path:
                    logger.error("LibreOffice (soffice) not found in PATH")
                    raise RuntimeError("LibreOffice (soffice) not found in PATH")
                soffice_path = str(Path(raw_cmd_path).resolve())

                # 🛡️ Security Fix: Use absolute paths to prevent argument injection
                res_pdf = subprocess.run(
                    [
                        soffice_path,
                        f"-env:UserInstallation=file://{tmp_dir.resolve()}",
                        "--headless",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(tmp_dir.resolve()),
                        str(temp_input.resolve()),
                    ],
                    capture_output=True,
                    text=True,
                    errors="replace",
                    timeout=600,
                )

                if res_pdf.returncode != 0:
                    logger.error(f"LibreOffice failed: {res_pdf.stderr}")
                    raise RuntimeError(f"LibreOffice conversion failed: {res_pdf.stderr}")

                # 生成されたPDFを固定の名前にリネームして管理しやすくする
                if expected_pdf.exists():
                    shutil.move(str(expected_pdf), str(original_pdf))

                if not original_pdf.exists():
                    raise FileNotFoundError(f"PDF not found after conversion: {original_pdf}")

                # Step 2: PDF -> PNG (multi-page support)
                logger.info(f"Converting PDF to PNG (multi-page: {bool(sheet_name)})...")

                if sheet_name:
                    # シートごとのマルチページ抽出
                    return self._convert_pdf_to_multi_png(original_pdf, output_prefix)
                else:
                    # 全体概要用の単一ファイル抽出 (後方互換性)
                    output_png = output_prefix.with_suffix(".png")
                    self._convert_pdf_to_png(original_pdf, output_png)
                    return output_png
            finally:
                # Cleanup inside the context manager while tmp_dir still exists
                if original_pdf.exists():
                    original_pdf.unlink()
                if temp_input.exists():
                    temp_input.unlink()

    def _convert_pdf_to_multi_png(self, pdf_path: Path, output_prefix: Path) -> List[Path]:
        """PDFの全ページをPNGに変換する"""
        raw_cmd_path = shutil.which("pdftocairo")
        if not raw_cmd_path:
            # pdftocairo がない場合は fitz を使用 (フォールバック)
            return self._try_fitz_multi(pdf_path, output_prefix)

        try:
            # 🛡️ Security Fix: Use absolute paths to prevent argument injection
            # pdftocairo を使用して全ページを連番で出力
            subprocess.run(
                [
                    str(Path(raw_cmd_path).resolve()),
                    "-png",
                    str(pdf_path.resolve()),
                    str(output_prefix.resolve()),
                ],
                check=True,
                capture_output=True,
                text=True,
                errors="replace",
                timeout=300,
            )

            # 生成されたファイルを収集 (prefix-1.png, prefix-2.png, ...)
            pngs = sorted(
                list(self.output_dir.glob(f"{output_prefix.name}-*.png")), key=lambda p: int(p.stem.split("-")[-1])
            )
            if pngs:
                logger.info(f"Generated {len(pngs)} PNGs using pdftocairo")
                return pngs
        except Exception as e:
            logger.warning(f"pdftocairo multi-page failed: {e}")

        return self._try_fitz_multi(pdf_path, output_prefix)

    def _try_fitz_multi(self, pdf_path: Path, output_prefix: Path) -> List[Path]:
        """PyMuPDFを使用して全ページをPNGに変換する"""
        try:
            import fitz

            doc = fitz.open(str(pdf_path.resolve()))
            pngs = []
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                out_path = self.output_dir / f"{output_prefix.name}-{i + 1}.png"
                pix.save(str(out_path.resolve()))
                pngs.append(out_path)
            doc.close()
            if pngs:
                logger.info(f"Generated {len(pngs)} PNGs using fitz")
                return pngs
        except Exception as e:
            logger.error(f"Multi-page conversion failed: {e}")

        raise RuntimeError("Multi-page PDF to PNG conversion failed")

    def _convert_pdf_to_png(self, pdf_path: Path, output_png: Path) -> None:
        """PDFをPNGに変換する (複数の方法を試行するフォールバックチェーン)"""

        # 1. pdftocairo (Primary)
        if self._try_pdftocairo(pdf_path, output_png):
            return

        # 2. PyMuPDF (fitz)
        if self._try_fitz(pdf_path, output_png):
            return

        # 3. ImageMagick (magick or convert)
        if self._try_imagemagick(pdf_path, output_png):
            return

        raise RuntimeError("All PDF to PNG conversion methods failed")

    def _try_pdftocairo(self, pdf_path: Path, output_png: Path) -> bool:
        """pdftocairoを使用してPDFをPNGに変換する"""
        raw_cmd_path = shutil.which("pdftocairo")
        if not raw_cmd_path:
            logger.warning("pdftocairo not found in PATH")
            return False

        try:
            # 🛡️ Security Fix: Use absolute paths to prevent argument injection and CWE-426
            res = subprocess.run(
                [
                    str(Path(raw_cmd_path).resolve()),
                    "-png",
                    "-singlefile",
                    str(pdf_path.resolve()),
                    str(output_png.with_suffix("").resolve()),
                ],
                capture_output=True,
                text=True,
                errors="replace",
                timeout=300,
            )

            if res.returncode == 0 and output_png.exists():
                logger.info("Converted PDF to PNG using pdftocairo")
                return True
            logger.warning(f"pdftocairo failed: {res.stderr}")
        except (subprocess.SubprocessError, OSError) as e:
            logger.warning(f"pdftocairo failed: {e}")
        return False

    def _try_fitz(self, pdf_path: Path, output_png: Path) -> bool:
        """PyMuPDF (fitz) を使用してPDFをPNGに変換する"""
        try:
            import fitz

            doc = fitz.open(str(pdf_path.resolve()))
            page = doc.load_page(0)
            # 2.0x zoom for better quality
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            pix.save(str(output_png.resolve()))
            doc.close()
            if output_png.exists():
                logger.info("Converted PDF to PNG using PyMuPDF (fitz)")
                return True
        except ImportError:
            logger.warning("PyMuPDF (fitz) not installed")
        except Exception as e:
            logger.warning(f"PyMuPDF conversion failed: {e}")
        return False

    def _try_imagemagick(self, pdf_path: Path, output_png: Path) -> bool:
        """ImageMagick (magick or convert) を使用してPDFをPNGに変換する"""
        for cmd_name in ["magick", "convert"]:
            raw_cmd_path = shutil.which(cmd_name)
            if not raw_cmd_path:
                continue
            try:
                # 🛡️ Security Fix: Use absolute paths to prevent argument injection
                # magick [input] [output] or convert [input] [output]
                # For PDF to PNG with ImageMagick, [0] specifies the first page
                res = subprocess.run(
                    [
                        str(Path(raw_cmd_path).resolve()),
                        "-density",
                        str(self.dpi),
                        f"{str(pdf_path.resolve())}[0]",
                        str(output_png.resolve()),
                    ],
                    capture_output=True,
                    text=True,
                    errors="replace",
                    timeout=300,
                )

                if res.returncode == 0 and output_png.exists():
                    logger.info(f"Converted PDF to PNG using ImageMagick ({cmd_name})")
                    return True
            except (subprocess.SubprocessError, OSError) as e:
                logger.warning(f"ImageMagick ({cmd_name}) failed: {e}")
                continue
        return False
