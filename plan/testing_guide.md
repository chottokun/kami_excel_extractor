# テスト実行ガイド

本プロジェクトでは、LibreOffice などの OS 依存ライブラリを使用するため、Docker コンテナ内でのテスト実行を推奨します。

## 1. コンテナ内での全テスト実行

コンテナが起動している状態で以下のコマンドを実行します。

```bash
docker exec pipeline-worker-lib uv run pytest tests/
```

### 特定のテストファイルのみ実行する場合
```bash
docker exec pipeline-worker-lib uv run pytest tests/test_core.py
```

## 2. ホスト側から最新のテストファイルを同期する

開発中にホスト側でテストコードを修正した場合、ボリュームマウントの設定によってはコンテナ内に反映されないことがあります（`docker-compose.yml` で `./tests` がマウントされていない場合など）。その場合は、以下のコマンドで同期してから実行してください。

```bash
docker cp tests/. pipeline-worker-lib:/app/tests/
docker exec pipeline-worker-lib uv run pytest tests/
```

## 3. 実データを用いた統合テスト

実際の Gemini API キーを使用して `complex_report.xlsx` などを処理するテストを行う場合は、`.env` に有効な `GEMINI_API_KEY` を設定した上で実行してください。

```bash
# コンテナ内での実データ検証スクリプト実行例
docker exec pipeline-worker-lib uv run python /app/src/main.py
```

## 4. テスト実装の注意点

- **非同期処理**: コアロジックが `asyncio` 化されているため、テストは `pytest-asyncio` を使用し、非同期メソッドを直接 `await` するように実装してください。
- **同期ラッパー**: `asyncio.run` を含む同期ラッパーメソッドを `pytest-asyncio` のループ内で呼び出すと `RuntimeError` が発生するため避けてください。
- **画像正規化**: 抽出した画像は必ず Pillow で PNG に正規化されることを確認してください。
- **JSON シリアライズ**: LLM や YAML から得られるデータに `date` や `datetime` オブジェクトが含まれる場合、最終出力前に文字列（ISO形式）へ変換されることを確認してください。
