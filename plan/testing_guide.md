# テスト実行ガイド

本プロジェクトでは、LibreOffice などの OS 依存ライブラリを使用するため、Docker コンテナ内でのテスト実行を推奨します。

## 1. コンテナ内での全テスト実行

コンテナが起動している状態で以下のコマンドを実行します。

```bash
# 全テストの実行
PYTHONPATH=src uv run pytest tests/
```

### 特定のテストファイルのみ実行する場合
```bash
PYTHONPATH=src uv run pytest tests/test_core.py
```

## 2. テストスイートの構成 (Test Map)

各テストファイルは特定の目的を持って設計されています。

| カテゴリ | テストファイル | 役割 |
| :--- | :--- | :--- |
| **コアロジック** | `test_core.py` | `KamiExcelExtractor` の非同期抽出フロー全体の検証。 |
| | `test_core_rag_chunks.py` | RAG 用チャンク生成と Markdown 変換の統合検証。 |
| **コンポーネント** | `test_extractor.py` | `MetadataExtractor` による HTML テーブル生成の検証。 |
| | `test_extractor_media.py` | 画像抽出、座標マッピング、Pillow 正規化の検証。 |
| | `test_converter.py` | `ExcelConverter` による PDF/PNG 変換とタイムアウトの検証。 |
| | `test_rag_converter.py` | `JsonToMarkdownConverter` の変換ロジック検証。 |
| | `test_document_generator.py` | Markdown から PDF を生成するレポート機能の検証。 |
| **特殊機能** | `test_simple_table.py` | LLM を介さない「物理抽出モード」の判定と抽出の検証。 |
| | `test_performance_optimization.py` | `iter_rows` や Simple Table による速度向上の検証。 |
| **非機能要件** | `test_security_remediation.py` | パス安全性、一時ディレクトリ管理、タイムアウト強制の検証。 |
| | `test_utils.py` | `secure_filename` 等のユーティリティ関数の検証。 |
| **外部連携** | `test_litellm.py` | `LiteLLM` との非同期通信およびモックの検証。 |

## 3. 実データを用いた統合テスト

実際の LLM API キー（Gemini 等）を使用して `complex_report.xlsx` などを処理するテストを行う場合は、`.env` に有効な `LLM_API_KEY` を設定した上で実行してください。

```bash
# 実データ検証スクリプトの実行例
PYTHONPATH=src uv run python src/main.py
```

## 4. テスト実装の重要事項

- **非同期処理**: コアロジックが `asyncio` 化されているため、テストは `pytest-asyncio` を使用し、非同期メソッドを直接 `await` してください。
- **モックの適用**: `subprocess.run` や `litellm.acompletion` は必ずモックし、環境依存や API コストの発生を抑えてください。
- **ファイルパス**: テスト内では `tmp_path` フィクスチャを使用し、実行環境を汚染しないようにしてください。
