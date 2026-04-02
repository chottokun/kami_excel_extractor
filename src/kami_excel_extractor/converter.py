import subprocess
import tempfile
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)

    def convert(self, input_file: Path) -> Path:
        output_png = self.output_dir / f"{input_file.stem}.png"

        # 入力ファイルの存在確認
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        with tempfile.TemporaryDirectory(prefix="lo_profile_") as tmp_dir:
            user_installation = f"file://{tmp_dir}"

            # Step 1: Excel -> PDF
            logger.info(f"Converting {input_file.name} to PDF...")
            # 🔒 Security Fix: Use absolute paths to prevent argument injection
            res_pdf = subprocess.run([
                "soffice", f"-env:UserInstallation={user_installation}",
                "--headless", "--convert-to", "pdf",
                "--outdir", str(self.output_dir.resolve()), str(input_file.resolve())
            ], capture_output=True, text=True, timeout=600)

            if res_pdf.returncode != 0:
                logger.error(f"LibreOffice failed: {res_pdf.stderr}")
                raise RuntimeError(f"LibreOffice conversion failed: {res_pdf.stderr}")

            original_pdf = self.output_dir / f"{input_file.stem}.pdf"
            if not original_pdf.exists():
                raise FileNotFoundError(f"PDF not found after conversion: {original_pdf}")

            # Step 2: PDF -> PNG
            logger.info(f"Converting PDF to PNG...")
            # 🔒 Security Fix: Use absolute paths to prevent argument injection
            res_png = subprocess.run([
                "pdftocairo", "-png", "-singlefile",
                str(original_pdf.resolve()), str((self.output_dir / input_file.stem).resolve())
            ], capture_output=True, text=True, timeout=300)

            if res_png.returncode != 0:
                logger.error(f"pdftocairo failed: {res_png.stderr}")
                if original_pdf.exists():
                    original_pdf.unlink()
                raise RuntimeError(f"pdftocairo conversion failed: {res_png.stderr}")

            if original_pdf.exists():
                original_pdf.unlink()

            return output_png
