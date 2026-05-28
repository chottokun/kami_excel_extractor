# Kami Excel Extractor

「方眼紙エクセル」や「非構造化された業務報告書」から、VLM/LLMを用いて構造化データ（JSON/Markdown）を抽出するための処理パイプラインです。
セルの視覚情報（罫線やフォントスタイル）と論理構造（計算式や表示形式）を解析し、VLM/LLMにコンテキストとして与えることで高精度な抽出を行います。

## 主な機能・仕様

- **Pagination-Aware Extraction**: 長尺のシートや複数シートを正確に認識するため、シートごとに隔離してマルチページ画像を生成し、VLMに提供。
- **High-Performance Caching**: OpenPyXLによるExcel解析結果や、VLM/LLMの応答をSQLiteにキャッシュし、同一ファイルの再処理時にAPI呼び出しをスキップして処理を効率化。
- **Style-Aware Extraction**: セルの罫線、背景色、フォントスタイルを CSS 形式で AI に伝達し、表の境界や見出しの認識を補助。
- **Logic-Aware Injection**: セルの「計算式」と「表示形式」を解析。合計値の特定や単位（円, %, 日付）の誤認を防止。
- **Visual Intelligence**: シート内の図表やグラフを VLM で自動解析し、数値データとして HTML 内に動的に埋め込み。
- **Semantic Cleaning**: レイアウト目的の不自然な空白（例: 「氏　名」）を自動除去し、正確な単語として復元。
- **Docker環境の提供**: LibreOffice, Poppler, ImageMagick を完備したコンテナ環境を提供し、ローカル環境の依存関係に左右されずに実行可能。
- **RAG & Search 最適化**: 構造化 JSON に加え、RAG 検索に適合する Markdown チャンクを同時に生成。
- **品質管理**: 型ヒントの適用、詳細なドキュメント整備、および 400 件以上のユニットテストによる堅牢な動作検証。

## ドキュメント

- **[ユーザマニュアル (USAGE.md)](USAGE.md)**: セットアップから CLI/API の詳細な使い方まで。
- **[抽出エンジン技術解説 (docs/extraction_engine.md)](docs/extraction_engine.md)**: 5つのレイヤーによる解析アルゴリズムの解説。
- **[改善提案ロードマップ (docs/improvement_proposals.md)](docs/improvement_proposals.md)**: 今後の精度向上に向けた考察。

## クイックスタート (Docker)

Docker Compose を使用した実行手順です。

```bash
# 1. 起動（data/input ディレクトリの監視を開始）
docker compose up -d --build

# 2. CLI による単発実行
docker compose run --rm cli sample.xlsx --include-logic --visual-summaries
```

## 基本的な使い方 (Local CLI)

```bash
# 高精度フルオプション解析
uv run python -m kami_excel_extractor.cli report.xlsx \
    --include-logic \
    --visual-summaries \
    --dpi 300
```

## セキュリティ対策

実務での利用を考慮し、以下のセキュリティ対策を実施しています：
- **CWE-426 (Untrusted Search Path) 対策**: 外部コマンドの絶対パス強制解決。
- **XSS 防御**: `quote=True` を含む HTML エスケープの強制適用。
- **DoS 対策**: 画像のピクセル数・ファイルサイズ制限による「画像爆弾（Image Bomb）」への耐性。
- **Path Traversal 防止**: 出力ファイル名のサニタイズ処理。

## ライセンス
MIT License

