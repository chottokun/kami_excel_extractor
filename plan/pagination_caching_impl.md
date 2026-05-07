# ページネーションとキャッシュ強化の実装計画

## 1. 目的
- **ページネーション**: 長尺シートや複数シートに跨る視覚情報を正確にVLMへ提供し、推論精度を向上させる。
- **キャッシュ強化**: 重いExcelパース（OpenPyXL）と画像生成（LibreOffice）をスキップ可能にし、開発効率と実行速度を劇的に改善する。

## 2. 実装詳細

### A. ページネーション (Multi-page Visual Context)
1.  **シートの隔離 (Isolation)**:
    - `ExcelConverter.convert()` を拡張し、特定のシートのみを抽出した一時的なExcelファイルを作成する機能を追加。
    - `openpyxl` を使用し、対象シート以外を削除して保存。
2.  **マルチページPDF/PNG生成**:
    - LibreOfficeで隔離されたExcelをPDF化。
    - `pdftocairo` 等で `-singlefile` を外して全ページをPNG出力。
3.  **VLMへの投入**:
    - `KamiExcelExtractor` で複数のBase64画像を収集し、`messages` に配列として投入。

### B. キャッシュ強化 (Multi-layer Caching)
1.  **Raw Extraction Cache**:
    - `MetadataExtractor.extract()` の戻り値（HTML, セル情報）を、ファイルのハッシュをキーにSQLiteに保存。
    - 2回目以降はOpenPyXLの読み込みを完全にスキップ。
2.  **Image Generation Cache**:
    - 生成されたPNG群のパスまたはBase64をキャッシュ。
    - ファイルハッシュとシート名が一致すれば、LibreOfficeの起動をスキップ。

## 3. 作業フェーズ

### フェーズ1: キャッシュ基盤の拡張
- `CacheManager` に `raw_extraction` テーブルを追加。
- `KamiExcelExtractor` でパース結果のキャッシュ読み書きを実装。

### フェーズ2: ページネーションの実装 (Converter)
- `ExcelConverter` に `sheet_name` 指定オプションを追加。
- シート隔離ロジックと、複数PNG返却ロジックの実装。

### フェーズ3: VLMオーケストレーションの更新 (Core)
- `aextract_structured_data` 内のループを調整し、各シートごとに画像生成を呼び出す。
- 複数画像をプロンプトに含めるように `_build_sheet_messages` を修正。

## 4. 検証計画
1.  **ユニットテスト**:
    - `test_converter.py` に複数ページ出力のテストを追加。
    - `test_cache_manager.py` にRaw Extractionキャッシュのテストを追加。
2.  **実データ検証**:
    - 複数ページに跨る `complex_report.xlsx` で全ページが画像化され、解析されることを確認。
3.  **目視確認**:
    - `output/media/` に各シート・各ページの画像が正しく生成されているか確認。
4.  **パフォーマンス測定**:
    - キャッシュ有効時の実行時間を計測し、高速化を実証。
