# Kami Excel Extractor

「方眼紙エクセル」や「複雑な業務報告書」を、視覚情報（画像）と論理構造（XMLメタデータ）の両面から解析し、VLM (Vision Language Model) を用いて高精度に構造化データ (JSON) へ変換するパイプラインです。

## 🌟 特徴

- **方眼紙エクセル対応**: セル結合や細かいグリッドを論理的に解釈し、ラベルと値のペアを正確に特定。
- **マルチモーダル解析**: LibreOffice による画像レンダリングと `openpyxl` による構造抽出を統合。
- **物理抽出最適化**: 単純な表構造を自動検知。LLMを使わず直接抽出することで、100%の精度確保とコスト削減を両立。
- **非同期・並列処理**: `asyncio` による複数シート・画像の並列解析。GPUリソースに応じた速度向上。
- **RAG最適化出力**: RAG検索精度を向上させる「Key: Value」形式のMarkdown出力に対応。
- **多形式アウトプット**: 解析結果を JSON, YAML, Markdown形式で同時出力。
- **高度な互換性**: 
    - 抽出画像を Pillow で自動正規化（PNG変換）し、VLMの読み取りエラーを防止。
    - PDFからPNGへの変換において、**3層のフォールバックチェーン**（pdftocairo, PyMuPDF, ImageMagick）を搭載。
- **セキュアな設計**: 徹底した HTML エスケープと引数注入対策、`tempfile` による安全な一時管理。

## 🚀 セットアップ

### 1. 依存パッケージ（Linux）
PDF変換や画像処理のために、以下のツールがインストールされていることが推奨されます。

```bash
# Ubuntu/Debian 例
sudo apt-get install libreoffice poppler-utils imagemagick
```

### 2. 環境設定
`.env` ファイルに LLM の提供プロバイダーに合わせた設定を行ってください。

```env
# LiteLLM 形式のモデル名 (openai/gpt-4o, gemini/gemini-1.5-flash, ollama/qwen3.5:4b 等)
LLM_MODEL=gemini/gemini-1.5-flash
LLM_API_KEY=your_api_key_here

# 必要に応じてベース URL やタイムアウト、RPM 制限を指定可能
# LLM_BASE_URL=http://localhost:11434
# LLM_TIMEOUT=600   # デフォルト 600秒(10分)
# LLM_RPM_LIMIT=15  # 1分あたりの最大リクエスト数
```

### 3. 起動
Docker Compose を使用して、監視パイプラインを起動します。

```bash
docker compose up -d --build
```

## 🛠️ 使い方 (CLI)

特定のファイルに対して手動で解析を実行できます。

```bash
# 基本的な実行
uv run python -m kami_excel_extractor.cli sample.xlsx --model gemini/gemini-1.5-flash

# RAG用Markdownチャンクも同時に生成
uv run python -m kami_excel_extractor.cli sample.xlsx --rag

# 画像解析を無効化（テキスト抽出のみ）
uv run python -m kami_excel_extractor.cli sample.xlsx --no-vision

# 詳細ログを表示
uv run python -m kami_excel_extractor.cli sample.xlsx --verbose
```

## 📦 ライブラリとしての利用

Python コード内から直接呼び出すことも可能です。

```python
import asyncio
from pathlib import Path
from kami_excel_extractor import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions

async def main():
    extractor = KamiExcelExtractor(output_dir="output")
    
    options = ExtractionOptions(
        model="gemini/gemini-1.5-flash",
        include_visual_summaries=True
    )
    
    result = await extractor.aextract_structured_data("report.xlsx", options=options)
    print(result)

asyncio.run(main())
```

## 🛡️ セキュリティと制限

### セキュリティ対策
- **XSS 対策**: 生成されるレポート内の全外部入力に対し HTML エスケープを徹底。
- **引数注入対策**: 外部コマンド（soffice, pdftocairo等）実行時の絶対パス解決 (CWE-426対策)。
- **パス・トラバーサル防止**: レポート出力名のサニタイズ (`secure_filename`)。
- **堅牢性**: 破損画像や推論エラー発生時も、プロセスを停止させずスキップ・記録。

### 既知の制限
- **特殊な埋め込み画像**: Excel内の「図形（Shape）」や「メタファイル（EMF/WMF）」、zlib圧縮された内部形式での画像埋め込みは現在抽出対象外（スキップ）となります。将来的に `pyvips` 等によるサポートを検討中です。

## 📜 ライセンス
[MIT License](LICENSE)
