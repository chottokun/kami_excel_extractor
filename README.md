# Kami Excel Extractor

「方眼紙エクセル」や「複雑な業務報告書」を、視覚情報（画像）と論理構造（XMLメタデータ）の両面から解析し、VLM (Vision Language Model) を用いて高精度に構造化データ (JSON) へ変換するパイプラインです。

## 🌟 特徴

- **方眼紙エクセル対応**: セル結合や細かいグリッドを論理的に解釈し、ラベルと値のペアを正確に特定。
- **マルチモーダル解析**: LibreOffice による画像レンダリングと `openpyxl` による構造抽出を統合。
- **メディア抽出**: Excel内に埋め込まれた現場写真や図解を物理ファイルとして抽出し、テキストと紐付け。
- **複数シート統合**: 複数のワークシートに分散した情報を一つの論理的なJSONに集約。
- **ローカル完結型変換**: 外部APIを使わずに、コンテナ内の LibreOffice Headless で安全に画像を生成。
- **再利用可能なライブラリ構成**: コアロジックを `kami_excel_extractor` パッケージとしてモジュール化。

## 🏗 アーキテクチャ

1.  **Rendering**: LibreOffice Headless (PDF) -> Poppler (PNG)
2.  **Extraction**: openpyxl による座標・スタイル・結合セル・画像データの抽出
3.  **Reasoning**: Gemini 1.5/2.0/2.5 Flash によるマルチモーダル推論
4.  **Output**: Pydantic スキーマ等に準拠可能な構造化 JSON

## 🚀 セットアップ

### 1. 環境設定
`.env` ファイルを作成し、Gemini APIキーを設定してください。

```env
GEMINI_API_KEY=your_api_key_here
GEMINI_MODEL=models/gemini-1.5-flash
```

### 2. 起動
Docker Compose を使用して、監視パイプラインを起動します。

```bash
docker compose up -d --build
```

`data/input/` に Excel ファイルを配置すると、自動的に処理が開始され、`data/output/` に結果が出力されます。

## 🧪 テストの実行

コンテナ内で `pytest` を実行して、画像変換や抽出ロジックの健全性を確認できます。

```bash
docker compose run --rm pipeline-worker uv run pytest tests/test_core.py
```

## 📦 ライブラリとしての利用

本プロジェクトのコア機能はモジュール化されており、自身のPythonプロジェクトから呼び出すことが可能です。

```python
from kami_excel_extractor import KamiExcelExtractor

extractor = KamiExcelExtractor(api_key="your_key", output_dir="output")
result = extractor.extract_structured_data("report.xlsx")
print(result)
```

## 📂 ディレクトリ構造

- `src/kami_excel_extractor/`: コアライブラリ（変換、抽出、統合）
- `src/main.py`: ファイル監視パイプラインの実行エントリーポイント
- `data/input/`: 処理待ちExcelファイルの投入ディレクトリ
- `data/output/`: 構造化JSONおよび抽出された写真の出力先
- `tests/`: pytest による単体テストおよび利用例テスト
