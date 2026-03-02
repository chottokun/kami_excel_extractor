import subprocess
import uuid
import os
import shutil
from pathlib import Path

def convert_to_png(input_file: Path, output_dir: Path) -> Path:
    """
    Excel -> PDF -> PNG の2段階変換を行い、ヘッドレス環境での安定性を確保する。
    """
    profile_id = uuid.uuid4().hex
    user_installation = f"file:///tmp/lo_profile_{profile_id}"
    output_png = output_dir / f"{input_file.stem}.png"
    
    # 出力先ディレクトリを確保
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Step 1: Excel -> PDF
        cmd_pdf = [
            "soffice",
            f"-env:UserInstallation={user_installation}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(input_file)
        ]
        print(f"[Step 1] Executing LibreOffice: {' '.join(cmd_pdf)}")
        res_pdf = subprocess.run(cmd_pdf, capture_output=True, text=True)
        print(f"[Step 1] PDF Output: {res_pdf.stdout} {res_pdf.stderr}")
        
        # 生成されたPDFを特定
        original_pdf = output_dir / f"{input_file.stem}.pdf"
        if not original_pdf.exists():
             raise FileNotFoundError(f"PDF creation failed for {input_file.name} at {original_pdf}")
        
        # Step 2: PDF -> PNG (pdftocairoを使用)
        cmd_png = [
            "pdftocairo",
            "-png",
            "-singlefile",
            str(original_pdf),
            str(output_dir / input_file.stem)
        ]
        print(f"[Step 2] Executing pdftocairo: {' '.join(cmd_png)}")
        res_png = subprocess.run(cmd_png, capture_output=True, text=True)
        print(f"[Step 2] PNG Output: {res_png.stdout} {res_png.stderr}")
        
        # 不要なPDFの削除
        original_pdf.unlink()
        
        if not output_png.exists():
            raise FileNotFoundError(f"PNG conversion failed: {output_png} not created.")
            
        print(f"[Success] Created: {output_png}")
        return output_png

    except Exception as e:
        print(f"[Error] Conversion process failed: {e}")
        raise
    finally:
        # 一時プロファイルの削除
        shutil.rmtree(f"/tmp/lo_profile_{profile_id}", ignore_errors=True)
