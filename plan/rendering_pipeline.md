# **ローカル画像レンダリング・パイプライン詳細**

LibreOffice Headless を利用し、Excelファイルを視覚情報（PNG）に変換するプロセスを管理する。

## **1. プロセス隔離と並行処理**

LibreOffice は単一ユーザープロファイルでの同時実行を許可しないため、各プロセスに動的なプロファイルディレクトリを割り当てる。

### **実装ロジック (Python)**
```python
import subprocess
import uuid
import os
from pathlib import Path

def convert_excel_to_png(input_path: Path, output_dir: Path):
    # プロセスごとにユニークなプロファイルを生成
    profile_id = uuid.uuid4().hex
    user_installation = f"file:///tmp/lo_profile_{profile_id}"
    
    cmd = [
        "soffice",
        f"-env:UserInstallation={user_installation}",
        "--headless",
        "--convert-to", "png",
        "--outdir", str(output_dir),
        str(input_path)
    ]
    
    try:
        subprocess.run(cmd, check=True, capture_output=True)
    finally:
        # プロファイルディレクトリのクリーンアップ（推奨）
        os.system(f"rm -rf /tmp/lo_profile_{profile_id}")
```

## **2. アーティファクト（ヘッダー・フッター）の除去**

VLMに対するノイズを最小限にするため、以下の設定を施したテンプレート（`templates/clean_layout.ots`）を適用する。

### **テンプレートに含める設定項目**
- **余白 (Margins):** 全方位 0mm
- **ヘッダー/フッター:** 無効 (Header/Footer On = False)
- **枠線表示:** セル枠線を含めて出力するか、データのみにするかの選択。

### **変換コマンドの改良**
テンプレートを反映させた変換を行う。
```bash
soffice --headless --infilter="Calc MS Excel 2007 XML" --convert-to png --outdir output_dir input.xlsx
```
※ コマンドラインからのテンプレート強制適用が困難な場合は、Python側で `openpyxl` を用いて、変換前にワークブック自体の `HeaderFooter` オブジェクトを初期化する前処理を挟む。

## **3. 課題と対策**

| 課題 | 対策 |
| :--- | :--- |
| **巨大なシート** | 出力が複数ページ（複数画像）に分かれる。Worker側で画像を結合するか、全画像をシーケンスとしてVLMに送る必要がある。 |
| **メモリ消費** | LibreOfficeプロセスはメモリを大量に消費するため、`ProcessPoolExecutor` の最大ワーカー数をコンテナのメモリ量に合わせて制限する。 |
| **サイレントフェイラー** | `subprocess.run` の終了コードを確認し、エラー時はリトライまたはログ記録を行う。 |
