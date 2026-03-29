# Kami Excel Extractor

「方眼紙エクセル」や「複雑な業務報告書」を、視覚情報（画像）と論理構造（XMLメタデータ）の両面から解析し、VLM (Vision Language Model) を用いて高精度に構造化データ (JSON) へ変換するパイプラインです。

## 🌟 特徴

- **方眼紙エクセル対応**: セル結合や細かいグリッドを論理的に解釈し、ラベルと値のペアを正確に特定。
- **マルチモーダル解析**: LibreOffice による画像レンダリングと `openpyxl` による構造抽出を統合。
- **物理抽出最適化 (NEW)**: 単純な表構造を自動検知。LLMを使わず直接抽出することで、100%の精度確保とコスト削減を両立。
- **非同期・並列処理 (NEW)**: `asyncio` による複数シート・画像の並列解析。実行時間を最大 90% 削減。
- **RAG最適化出力 (NEW)**: RAG検索精度を向上させる「Key: Value」形式のMarkdown出力に対応。
- **高度な互換性 (NEW)**: 
    - 抽出画像を Pillow で自動正規化（PNG変換）し、VLMの読み取りエラー（400）を防止。
    - Excel特有の日付データ（datetime）を自動的に文字列（ISO形式）へ変換し、JSON出力を保証。
- **セキュアな一時ファイル管理**: `tempfile.TemporaryDirectory` による安全なディレクトリ処理。

## 🚀 セットアップ

### 1. 環境設定
`.env` ファイルに LLM の提供プロバイダーに合わせた設定を行ってください。

```env
# LiteLLM 形式のモデル名 (openai/gpt-4o, gemini/gemini-1.5-flash, ollama/qwen3.5:4b 等)
LLM_MODEL=gemini/gemini-1.5-flash
LLM_API_KEY=your_api_key_here

# 必要に応じてベース URL やタイムアウト、RPM 制限を指定可能
# LLM_BASE_URL=http://localhost:11434
# LLM_TIMEOUT=1800  # 複雑な解析には 30分(1800s) 程度を推奨
# LLM_RPM_LIMIT=15
```

### 2. Ollama (ローカル LLM) の利用
Ollama を使用してローカル環境で抽出を行うことも可能です。詳細は [Ollama 利用ガイド](docs/ollama.md) を参照してください。
※ Docker コンテナ内からホスト側の Ollama に接続する場合は、`LLM_BASE_URL=http://172.17.0.1:11434` を指定してください。

### 3. 起動
Docker Compose を使用して、監視パイプラインを起動します。

```bash
docker compose up -d --build
```

## 🛡️ セキュリティと堅牢性

本ツールは業務利用を想定し、以下のセキュリティ対策を標準で備えています。

- **XSS 対策**: 生成される HTML/PDF レポート内のすべての外部入力（Excelの内容や画像属性）に対して HTML エスケープを徹底しています。
- **引数注入対策**: 外部コマンド（LibreOffice, pdftocairo）実行時、ファイルパスを絶対パスに正規化することで意図しない引数の注入を防止しています。
- **堅牢なエラーハンドリング**: 破損した画像や解析不可能なデータが含まれていても、プロセスを停止させずスキップ・記録する設計になっています。

## 🚀 CLI での利用

インストール後、`kami-excel` コマンドが利用可能になります。

```bash
# 基本的な利用 (Geminiを使用)
kami-excel report.xlsx

# Ollama を使用してテキストのみで抽出
kami-excel report.xlsx --model ollama/qwen3.5:4b --no-vision --base-url http://localhost:11434

# RAG用データの生成
kami-excel report.xlsx --rag --output-dir ./my_data
```

## 📦 ライブラリとしての利用

### 非同期呼び出し (推奨)
```python
import asyncio
from kami_excel_extractor import KamiExcelExtractor

async def main():
    # .env の設定が自動的に読み込まれます
    extractor = KamiExcelExtractor()
    
    # 画像概要の生成も含めて並列実行
    result = await extractor.aextract_structured_data("report.xlsx", include_visual_summaries=True)
    print(result)

asyncio.run(main())
```

### RAG用チャンク生成
```python
# list_format="kv" を指定することで検索に強い形式で出力
sheet_results, full_data = await extractor.aextract_rag_chunks("report.xlsx", list_format="kv")
```

## 🧪 テストの実行

```bash
# 全テストの実行
PYTHONPATH=src uv run pytest tests/
```

## 📂 ディレクトリ構造

- `src/kami_excel_extractor/`: コアライブラリ（完全パッケージ化）
- `data/input/`: 処理待ちExcelの投入先
- `data/output/`: 構造化データおよび抽出写真の出力先
- `tests/`: 徹底的な検証済みテストスイート（セキュリティ、非同期、物理抽出等）
