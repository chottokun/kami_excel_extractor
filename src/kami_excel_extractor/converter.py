import subprocess
import tempfile
import logging
import shutil
from pathlib import Path

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""
    
    def __init__(self, output_dir: Path, dpi: int = 150):
        self.output_dir = Path(output_dir).resolve()
        self.dpi = dpi

    def convert(self, input_file: Path) -> Path:
        input_file = input_file.resolve()
        output_png = self.output_dir / f"{input_file.stem}.png"
        original_pdf = self.output_dir / f"{input_file.stem}.pdf"

        # Cleanup existing targets to avoid permission/stale issues
        if output_png.exists():
            output_png.unlink()
        if original_pdf.exists():
            original_pdf.unlink()

        # 入力ファイルの存在確認
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        try:
            with tempfile.TemporaryDirectory(prefix="lo_profile_") as tmp_dir_str:
                tmp_dir = Path(tmp_dir_str).resolve()
                user_installation = f"file://{tmp_dir}"

                # Step 1: Excel -> PDF
                logger.info(f"Converting {input_file.name} to PDF...")
                
                # 🔒 Security Fix: Use absolute path for executable to prevent untrusted search path (CWE-426)
                raw_soffice_path = shutil.which("soffice")
                if not raw_soffice_path:
                    logger.error("LibreOffice (soffice) not found in PATH")
                    raise RuntimeError("LibreOffice (soffice) not found in PATH")
                soffice_path = str(Path(raw_soffice_path).resolve())

                # 🔒 Security Fix: Use absolute paths to prevent argument injection
                res_pdf = subprocess.run([
                    soffice_path, f"-env:UserInstallation={user_installation}",
                    "--headless", "--convert-to", "pdf",
                    "--outdir", str(self.output_dir.resolve()), str(input_file.resolve())
                ], capture_output=True, text=True, timeout=600)

                if res_pdf.returncode != 0:
                    logger.error(f"LibreOffice failed: {res_pdf.stderr}")
                    raise RuntimeError(f"LibreOffice conversion failed: {res_pdf.stderr}")

                if not original_pdf.exists():
                    raise FileNotFoundError(f"PDF not found after conversion: {original_pdf}")

                # Step 2: PDF -> PNG (with fallbacks)
                logger.info("Converting PDF to PNG...")
                self._convert_pdf_to_png(original_pdf, output_png)

                return output_png
        finally:
            if original_pdf.exists():
                original_pdf.unlink()

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
        raw_path = shutil.which("pdftocairo")
        if not raw_path:
            logger.warning("pdftocairo not found in PATH")
            return False

        try:
            # 🔒 Security Fix: Use absolute paths to prevent argument injection and CWE-426
            pdftocairo_path = str(Path(raw_path).resolve())
            res = subprocess.run([
                pdftocairo_path, "-png", "-singlefile",
                str(pdf_path.resolve()), str(output_png.with_suffix("").resolve())
            ], capture_output=True, text=True, timeout=300)

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
            raw_path = shutil.which(cmd_name)
            if not raw_path:
                continue
            try:
                # 🔒 Security Fix: Use absolute paths to prevent argument injection and CWE-426
                cmd_path = str(Path(raw_path).resolve())
                res = subprocess.run([
                    cmd_path, "-density", str(self.dpi),
                    f"{pdf_path.resolve()}[0]", str(output_png.resolve())
                ], capture_output=True, text=True, timeout=300)

                if res.returncode == 0 and output_png.exists():
                    logger.info(f"Converted PDF to PNG using ImageMagick ({cmd_name})")
                    return True
            except (subprocess.SubprocessError, OSError) as e:
                logger.warning(f"ImageMagick ({cmd_name}) failed: {e}")
                continue
        return False
