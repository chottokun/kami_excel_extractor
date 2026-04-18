# Kami Excel Extractor

「方眼紙エクセル」や「非構造化された業務報告書」を、人間のような知性で読み解く。
視覚情報（スタイル・画像）と論理構造（計算式・メタデータ）を統合し、VLM (Vision Language Model) を用いて高精度に構造化データ (JSON) へ変換する次世代抽出パイプラインです。

## 🌟 主な特徴

- **Style-Aware Extraction**: セルの罫線、背景色、フォントスタイルを CSS 形式で AI に伝達。表の境界や見出しを視覚的に理解。
- **Logic-Aware Injection**: セルの「計算式」と「表示形式」を解析。合計値の特定や単位（円, %, 日付）の誤認を防止。
- **Visual Intelligence**: シート内の図表やグラフを VLM で自動解析し、数値データとして HTML 内に動的に埋め込み。
- **Semantic Cleaning**: レイアウト目的の不自然な空白（例: 「氏　名」）を自動除去し、正確な単語として復元。
- **Docker-First 設計**: LibreOffice, Poppler, ImageMagick を完備したコンテナ環境で、依存関係なしに即座に実行可能。
- **RAG & Search 最適化**: 構造化 JSON に加え、RAG 検索精度を最大化する Markdown チャンクを同時生成。
- **プロフェッショナル品質**: 徹底した型ヒント、詳細なドキュメント、および 170 件以上の自動テストによる堅牢な信頼性。

## 📖 ドキュメント

詳細な情報は以下のドキュメントを参照してください：
- **[ユーザマニュアル (USAGE.md)](USAGE.md)**: セットアップから CLI/API の詳細な使い方まで。
- **[抽出エンジン技術解説 (docs/extraction_engine.md)](docs/extraction_engine.md)**: 5つのレイヤーによる高度な解析アルゴリズムの解説。
- **[改善提案ロードマップ (docs/improvement_proposals.md)](docs/improvement_proposals.md)**: 今後の精度向上に向けた考察。

## 🚀 クイックスタート (Docker)

最も推奨される実行方法は Docker Compose です。

```bash
# 1. 起動（data/input ディレクトリの監視を開始）
docker compose up -d --build

# 2. CLI による単発実行
docker compose run --rm cli sample.xlsx --include-logic --visual-summaries
```

## 🛠️ 基本的な使い方 (Local CLI)

```bash
# 高精度フルオプション解析
uv run python -m kami_excel_extractor.cli report.xlsx \
    --include-logic \
    --visual-summaries \
    --dpi 300
```

## 🛡️ セキュリティ

本プロジェクトは実務での利用を想定し、以下のセキュリティ対策を徹底しています：
- **CWE-426 (Untrusted Search Path) 対策**: 外部コマンドの絶対パス強制解決。
- **XSS 防御**: `quote=True` を含む厳格な HTML エスケープ。
- **DoS 対策**: 画像のピクセル数・ファイルサイズ制限による「画像爆弾」への耐性。
- **Path Traversal 防止**: 出力ファイル名の厳格なサニタイズ。

## 📜 ライセンス
MIT License
