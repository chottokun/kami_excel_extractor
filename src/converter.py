import subprocess
import uuid
import os
import shutil
from pathlib import Path

def convert_to_png(input_file: Path, output_dir: Path) -> Path:
    """
    LibreOffice Headlessを使用してExcelファイルをPNGに変換する。
    並行処理時の競合を避けるため、一意のUserInstallationプロファイルを使用する。
    """
    profile_id = uuid.uuid4().hex
    user_installation = f"file:///tmp/lo_profile_{profile_id}"
    
    # 出力先ディレクトリを確保
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # LibreOfficeコマンドの構築
    # ※方眼紙や複雑なレイアウトのため、まずはPDF経由ではなく直接PNG変換を試みる
    cmd = [
        "soffice",
        f"-env:UserInstallation={user_installation}",
        "--headless",
        "--convert-to", "png",
        "--outdir", str(output_dir),
        str(input_file)
    ]
    
    print(f"Converting {input_file.name} to PNG (Profile: {profile_id})...")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        # 生成されたファイル名を特定（拡張子がpngに変わる）
        output_file = output_dir / f"{input_file.stem}.png"
        if not output_file.exists():
            raise FileNotFoundError(f"LibreOffice conversion failed: {output_file} not created.")
        return output_file
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
