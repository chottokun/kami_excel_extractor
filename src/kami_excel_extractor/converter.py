import subprocess
import uuid
import shutil
from pathlib import Path

class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""
    
    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)

    def convert(self, input_file: Path) -> Path:
        profile_id = uuid.uuid4().hex
        user_installation = f"file:///tmp/lo_profile_{profile_id}"
        output_png = self.output_dir / f"{input_file.stem}.png"
        
        try:
            # Step 1: Excel -> PDF
            subprocess.run([
                "soffice", f"-env:UserInstallation={user_installation}",
                "--headless", "--convert-to", "pdf",
                "--outdir", str(self.output_dir), str(input_file)
            ], check=True, capture_output=True)
            
            original_pdf = self.output_dir / f"{input_file.stem}.pdf"
            
            # Step 2: PDF -> PNG
            subprocess.run([
                "pdftocairo", "-png", "-singlefile",
                str(original_pdf), str(self.output_dir / input_file.stem)
            ], check=True, capture_output=True)
            
            if original_pdf.exists():
                original_pdf.unlink()
                
            return output_png
        finally:
            shutil.rmtree(f"/tmp/lo_profile_{profile_id}", ignore_errors=True)
