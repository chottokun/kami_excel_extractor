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
    - Excel特有の日付データ（datetime）を自動的に文字列（ISO形式）へ変換。
- **セキュアな設計**: 徹底した HTML エスケープと引数注入対策、`tempfile` による安全な一時管理。

## 🚀 セットアップ

### 1. 環境設定
`.env` ファイルに LLM の提供プロバイダーに合わせた設定を行ってください。

```env
# LiteLLM 形式のモデル名 (openai/gpt-4o, gemini/gemini-1.5-flash, ollama/qwen3.5:4b 等)
LLM_MODEL=gemini/gemini-1.5-flash
LLM_API_KEY=your_api_key_here

# 必要に応じてベース URL やタイムアウト、RPM 制限を指定可能
# LLM_BASE_URL=http://localhost:11434
# LLM_TIMEOUT=1800  # 複雑な解析（方眼紙など）には 1800秒(30分) 以上を推奨
# LLM_RPM_LIMIT=15  # 1分あたりの最大リクエスト数
```

### 2. Ollama (ローカル LLM) の利用
Ollama を使用してローカル環境で抽出を行うことも可能です。
※ Linux Docker コンテナ内からホスト側の Ollama に接続する場合は、`LLM_BASE_URL=http://172.17.0.1:11434` を使用してください。詳細は [Ollama 利用ガイド](docs/ollama.md) を参照。

### 3. 起動
Docker Compose を使用して、監視パイプラインを起動します。

```bash
docker compose up -d --build
```

## 📂 ディレクトリ構成と成果物

- `src/kami_excel_extractor/`: コアライブラリ（完全パッケージ化）
- `data/input/`: 処理待ちExcelの投入先
- `data/output/`: 成果物の出力先。ファイルごとに以下の構成で出力されます。
    - `full_lib_result.json`: 全シートの統合抽出結果
    - `(シート名)_lib_result.json/yaml`: シートごとの構造化データ
    - `(シート名)_rag.md`: RAG/閲覧用の Markdown
    - **`(シート名)_report.pdf`**: セキュアに生成されたビジュアル報告書

## 🧪 テストと検証

```bash
# 全自動テストの実行 (コンテナ内推奨)
docker exec pipeline-worker-lib uv run pytest tests/

# 特定のファイルに対する手動解析 (CLI)
docker exec pipeline-worker-lib uv run python -m kami_excel_extractor.cli data/input/sample.xlsx --model ollama/qwen3.5:4b
```

## 🛡️ セキュリティ

本ツールは業務利用を想定し、以下の対策を標準で備えています。
- **XSS 対策**: 生成されるレポート内の全外部入力（Excel内容、画像属性等）に対し HTML エスケープを徹底。
- **引数注入対策**: 外部コマンド（LibreOffice等）実行時のパス正規化。
- **堅牢なエラーハンドリング**: 破損画像や推論エラー発生時も、プロセスを停止させずスキップ・記録。
