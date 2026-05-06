# Kami Excel Extractor ユーザマニュアル

本ツールは、複雑なレイアウト、結合セル、画像、および計算式を含む「神エクセル」から、高精度に構造化データ（JSON/Markdown）を抽出するための AI パワード・エンジンです。

## 1. 準備

### 1.1 前提条件
- **Python 3.10以上**
- **uv** (推奨されるパッケージマネージャー)
- **LibreOffice** (PDF変換および画像解析（Vision）機能を使用する場合に必要)
- **AI APIキー** (Gemini 1.5 Pro, GPT-4o, または Ollama 等)

### 1.2 インストール
```bash
# リポジトリのクローン
git clone <repository_url>
cd kami_excel_extractor

# 依存関係のインストール
uv sync
```

### 1.3 環境設定
`.env` ファイルを作成し、使用するモデルの API キーを設定します。
```env
GEMINI_API_KEY=your_key_here
LLM_MODEL=gemini-1.5-pro
```

## 2. 基本的な使い方 (CLI)

最もシンプルな実行方法は、エクセルファイルを指定して CLI を実行することです。

```bash
uv run python3 -m kami_excel_extractor.cli input.xlsx --output-dir ./results
```

### 主要なコマンドライン引数
- `--model`: 使用する AI モデルを指定 (例: `gemini-1.5-pro`)。
- `--rag`: 構造化 JSON に加え、RAG（検索拡張生成）に最適な Markdown チャンクを生成します。
- `--include-logic`: **[NEW]** エクセル内の計算式（例: `=SUM(A1:B1)`）と表示形式（円、%、日付）を抽出し、AI に伝えます。
- `--no-vision`: 画像解析を無効化し、テキストと構造情報のみで実行します。
- `--dpi`: 画像変換時の解像度を指定（デフォルト: 150）。

## 3. 高度な抽出機能

### 3.1 視覚スタイル認識 (Style-aware)
セルの罫線の太さや背景色を CSS 形式で AI に伝えます。

### 3.2 セマンティック・クリーニング (Semantic Cleaning)
不自然な空白を自動的に除去します。

### 3.3 Visual Intelligence (図表データ統合)
シート内の画像やグラフを VLM で解析し、内容を HTML テーブルの中に埋め込みます。

### 3.4 ロジック・インジェクション (Logic-aware)
セルの「計算式」を抽出し、集計構造の把握を助けます。

## 4. プログラマ向けの使い方 (Python API)

```python
import asyncio
from kami_excel_extractor import KamiExcelExtractor, ExtractionOptions

async def main():
    extractor = KamiExcelExtractor(output_dir="output")
    options = ExtractionOptions(
        model="gemini-1.5-pro",
        include_logic=True,
        include_visual_summaries=True,
        use_cache=True # 永続キャッシュを有効化
    )
    
    result = await extractor.aextract_structured_data("report.xlsx", options=options)
    print(result["sheets"]["Sheet1"])

if __name__ == "__main__":
    asyncio.run(main())
```

### 高度なオプション (`ExtractionOptions`)

| オプション | デフォルト | 説明 |
| :--- | :--- | :--- |
| `use_cache` | `True` | **[NEW]** 解析結果を SQLite にキャッシュし、2回目以降を高速化します。 |
| `max_file_size_mb` | `50` | **[NEW]** OOM 防止のため、指定サイズ以上のエクセルを拒否します。 |
| `include_logic` | `False` | 計算式や表示形式の抽出を有効にします。 |

## 5. 信頼性と堅牢性の機能 (重要)

### 5.1 指数関数的バックオフ (API リトライ)
API のレート制限 (`429`) やネットワークエラー時、自動で待機時間を増やしながら再試行します。

### 5.2 永続キャッシュの管理
解析結果は `output/.cache.db` に保存されます。画像の内容（ハッシュ）をキーにするため、同一ファイル名でも中身が変われば自動で再解析されます。
**キャッシュのクリア:**
```bash
rm output/.cache.db
```

### 5.3 巨大ファイルの保護
`max_file_size_mb` 設定により、メモリ不足によるシステムクラッシュを未然に防ぎます。

### 5.4 バランスパース (画像名の堅牢な解析)
画像ファイル名に括弧が含まれる場合（例: `data(1).png`）や、複雑にネストされている場合も正しく処理できます。

### 5.5 画像のメモリ枯渇DoS脆弱性の保護 (DoS Protection)
**[NEW]** 大容量の悪意ある画像（ZIP爆弾等）によるサーバーのメモリクラッシュ（OOM）を防ぐため、ストリームからの画像抽出時に 8KB チャンクでの読み込み制限を導入しました。`MAX_IMAGE_BYTES`（デフォルト 20MB）の上限に達した時点で自動的に処理をスキップ・アボートします。さらに、`getbuffer()` メモリ最適化により、バッファデータの余分なコピーを完全に排除しています。

### 5.6 メディア同期処理の O(1) 高速化
**[NEW]** 大量に画像や図表が埋め込まれた巨大なエクセルを並列抽出する際のボトルネックを解消するため、メタデータ同期処理に O(1) Lookup マップを導入しました。これにより、同期にかかる時間計算量が `O(N^2)` から `O(N)` に劇的に削減され、並列バッチ処理が極めて高速に完走します。

### 5.7 非同期PDF生成とスレッド再利用
**[NEW]** `DocumentGenerator` において、PDF生成の完全非同期関数 `agenerate_pdf` を新規実装しました。毎回新規スレッドを起動するオーバーヘッドを避けるため、永続スレッドプール（`ThreadPoolExecutor`）を再利用する設計を採用。画像解決の並列非同期化（`asyncio.gather`）と重い変換処理全体の非ブロック・スレッド退避により、APIサーバー等の高スループット環境でも完全にノンブロッキングで高速に動作します。
