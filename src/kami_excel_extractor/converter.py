import subprocess
import uuid
import shutil
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)

    def convert(self, input_file: Path) -> Path:
        profile_id = uuid.uuid4().hex
        user_installation = f"file:///tmp/lo_profile_{profile_id}"
        output_png = self.output_dir / f"{input_file.stem}.png"
        
        # 入力ファイルの存在確認
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        try:
            # Step 1: Excel -> PDF
            logger.info(f"Converting {input_file.name} to PDF...")
            res_pdf = subprocess.run([
                "soffice", f"-env:UserInstallation={user_installation}",
                "--headless", "--convert-to", "pdf",
                "--outdir", str(self.output_dir), str(input_file)
            ], capture_output=True, text=True)
            
            if res_pdf.returncode != 0:
                logger.error(f"LibreOffice failed: {res_pdf.stderr}")
                raise RuntimeError(f"LibreOffice conversion failed: {res_pdf.stderr}")
            
            original_pdf = self.output_dir / f"{input_file.stem}.pdf"
            if not original_pdf.exists():
                raise FileNotFoundError(f"PDF not found after conversion: {original_pdf}")
            
            # Step 2: PDF -> PNG
            logger.info(f"Converting PDF to PNG...")
            res_png = subprocess.run([
                "pdftocairo", "-png", "-singlefile",
                str(original_pdf), str(self.output_dir / input_file.stem)
            ], capture_output=True, text=True)
            
            if res_png.returncode != 0:
                logger.error(f"pdftocairo failed: {res_png.stderr}")
                raise RuntimeError(f"pdftocairo conversion failed: {res_png.stderr}")
            
            if original_pdf.exists():
                original_pdf.unlink()
                
            return output_png
        finally:
            shutil.rmtree(f"/tmp/lo_profile_{profile_id}", ignore_errors=True)
