# 網羅的動作確認手順 (Operational Check Guide)

本プロジェクトの全機能を網羅的にテストし、正常動作を確認するための手順書です。

## 1. 準備事項

- **環境**: Docker および Docker Compose がインストールされていること。
- **設定**: `.env` ファイルに有効な `GEMINI_API_KEY` (または各プロバイダーのキー) が設定されていること。
- **データ**: `complex_report.xlsx` などのテスト用 Excel ファイルがルートディレクトリに存在すること。

## 2. 動作確認手順

### ステップ 1: サービスの一括起動 (常駐監視モード)

常駐型の解析ワーカーを起動し、入力ディレクトリの監視を開始します。

```bash
# サービスのビルドと起動
docker compose up -d --build

# ログの監視
docker compose logs -f pipeline-worker
```

### ステップ 2: ファイルの投入と自動解析の確認

監視対象ディレクトリにファイルをコピーし、自動的に解析が始まることを確認します。

```bash
# 入力ディレクトリへのコピー
mkdir -p data/input
cp complex_report.xlsx data/input/

# ログで「Success: Outputs saved to ...」と表示されるのを待機
```

### ステップ 3: 出力結果の目視確認

`data/output` ディレクトリに生成されたファイルを確認します。

```bash
ls -R data/output/complex_report/
```

以下のファイルが生成されていることを確認してください：
- `full_lib_result.json`: 全シートの統合構造化データ
- `{SheetName}_lib_result.json`: シート個別の構造化データ
- `{SheetName}_lib_result.yaml`: LLM の生の応答テキスト
- `{SheetName}_rag.md`: RAG 用に最適化された Markdown
- `{SheetName}_rag_chunks.json`: Markdown を分割したチャンクデータ
- `{SheetName}_report.pdf`: 視覚的に再現された PDF レポート

### ステップ 4: CLI モードによる単発実行の確認

特定のオプションを指定して、オンデマンドで解析を実行します。

```bash
docker compose run --rm cli complex_report.xlsx \
    --include-logic \
    --visual-summaries \
    --output-dir /app/data/output
```

### ステップ 5: ユニットテストの実行

コンテナ環境内で全ての自動テストを実行し、環境依存の不具合がないか確認します。

```bash
docker compose run --rm cli uv run pytest tests/ -v --tb=short
```

## 3. トラブルシューティング

- **Permission Denied**: `data/output` ディレクトリの書き込み権限を確認してください。
- **PDF Generation Failed**: コンテナ内に `libreoffice` がインストールされているか確認してください（`Dockerfile.worker` でインストールされます）。
- **LLM Error**: `.env` の API キーが正しいか、モデル名が LiteLLM でサポートされている形式か確認してください。

---
最終更新日: 2026年5月8日
