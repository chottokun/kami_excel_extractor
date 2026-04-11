# テスト実行ガイド (完全版)

本プロジェクトは 140 件を超える網羅的なテストスイートを備え、単体テストから実データ End-to-End テスト、さらには高度なセキュリティ・並列制御の検証までを自動化しています。

## 1. 推奨実行環境

LibreOffice, Poppler, ImageMagick を含む Docker コンテナ内での実行を強く推奨します。

```bash
# 全テスト (141件) の一斉実行
docker exec pipeline-worker-lib uv run pytest
```

## 2. テストスイート詳細マップ (141 Cases)

| カテゴリ | 対象・目的 | 主なテストファイル |
| :--- | :--- | :--- |
| **End-to-End (コア)** | 非同期フロー、リトライ、抽出統合 | `test_core.py`, `test_main.py` |
| **セキュリティ (🔒)** | パストラバーサル, 引数注入, XSS, ImageBomb | `test_security_*.py`, `test_utils_edge_cases.py` |
| **並列制御 (⚡)** | Semaphore による流量制限, RPM 制御 | `test_core_concurrency.py` |
| **堅牢性 (🧪)** | YAML パース失敗、画像破損、モデル解決 | `test_core_yaml_failure.py`, `test_extractor_image_failures.py` |
| **RAG/変換** | Markdown 変換, 複雑なチャンク分割 | `test_rag_converter.py`, `test_rag_robustness.py` |
| **最適化** | シンプル表のバイパス, 正規化, キャッシュ | `test_performance_optimization.py`, `test_simple_table.py` |
| **変換エンジン** | PDF/PNG フォールバックチェーン | `test_converter.py` |
| **スキーマ** | Pydantic モデルによる構成・バリデーション | `test_schema_validation.py` |

## 3. 実データ検証 (End-to-End)

モックではない「真の動作」を確認するためのスクリプトが用意されています。

- **`run_container_verify.py`**: コンテナ内で 25 件の実 Excel ファイルを直列処理し、結果の整合性を検証します。
- **`run_exhaustive_test.py`**: ローカル環境で並列度を最大にして負荷耐性を検証します。

## 4. 品質保証のポイント

- **セキュリティ境界値**: `secure_filename` は NFKD 正規化と正規表現サニタイズを組み合わせ、日本語を維持しつつディレクトリ・トラバーサルを完全に遮断します。
- **非同期セーフ**: 全ての I/O 処理 (`to_thread`) とネットワークリクエスト (`Semaphore`) が適切に排他制御されていることを `test_core_concurrency` で保証しています。
- **LiteLLM 安全性**: v1.83.4 以降を使用し、サプライチェーン攻撃のリスクを排除した状態で全機能が動作することを確認済みです。
