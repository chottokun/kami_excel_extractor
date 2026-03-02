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
    temp_pdf = output_dir / f"{input_file.stem}_{profile_id}.pdf"
    output_png = output_dir / f"{input_file.stem}.png"
    
    # 出力先ディレクトリを確保
    output_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Step 1: Excel -> PDF (LibreOfficeはPDF変換が非常に安定している)
        cmd_pdf = [
            "soffice",
            f"-env:UserInstallation={user_installation}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(input_file)
        ]
        print(f"Converting {input_file.name} to PDF...")
        subprocess.run(cmd_pdf, check=True, capture_output=True)
        
        # 生成されたPDFを特定（sofficeは入力ファイル名に基づき作成する）
        original_pdf = output_dir / f"{input_file.stem}.pdf"
        if not original_pdf.exists():
             raise FileNotFoundError(f"PDF creation failed for {input_file.name}")
        
        # Step 2: PDF -> PNG (pdftocairoを使用)
        # -singlefile オプションで1つのファイルにまとめ、1ページ目のみを抽出
        cmd_png = [
            "pdftocairo",
            "-png",
            "-singlefile",
            str(original_pdf),
            str(output_dir / input_file.stem) # 拡張子は自動で付与される
        ]
        print(f"Converting PDF to PNG...")
        subprocess.run(cmd_png, check=True, capture_output=True)
        
        # 不要なPDFの削除
        original_pdf.unlink()
        
        if not output_png.exists():
            raise FileNotFoundError(f"PNG conversion failed: {output_png} not created.")
            
        return output_png

    except subprocess.CalledProcessError as e:
        print(f"Error during conversion: {e.stderr}")
        raise
    finally:
        # 一時プロファイルの削除
        shutil.rmtree(f"/tmp/lo_profile_{profile_id}", ignore_errors=True)

if __name__ == "__main__":
    # テスト実行用
    test_file = Path("sample_hoganshi.xlsx")
    if test_file.exists():
        out = convert_to_png(test_file, Path("data/output"))
        print(f"Success: {out}")
