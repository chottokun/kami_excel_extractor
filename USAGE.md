# Kami Excel Extractor ユーザマニュアル

本ツールは、複雑なレイアウト、結合セル、画像、および計算式を含む「神エクセル」から、高精度に構造化データ（JSON/Markdown）を抽出するための AI パワード・エンジンです。

## 1. 準備

### 1.1 前提条件
- **Python 3.10以上**
- **uv** (推奨されるパッケージマネージャー)
- **LibreOffice** (PDF変換および画像解析（Vision）機能を使用する場合に必要)
- **AI APIキー** (Gemini 1.5 Pro, GPT-4o, または Ollama 等)

### 1.2 インストール
```bash
# リポジトリのクローン
git clone <repository_url>
cd kami_excel_extractor

# 依存関係のインストール
uv sync
```

### 1.3 環境設定
`.env` ファイルを作成し、使用するモデルの API キーを設定します。
```env
GEMINI_API_KEY=your_key_here
LLM_MODEL=gemini-1.5-pro
```

## 2. 基本的な使い方 (CLI)

最もシンプルな実行方法は、エクセルファイルを指定して CLI を実行することです。

```bash
uv run python3 -m kami_excel_extractor.cli input.xlsx --output-dir ./results
```

### 主要なコマンドライン引数
- `--model`: 使用する AI モデルを指定 (例: `gemini-1.5-pro`)。
- `--rag`: 構造化 JSON に加え、RAG（検索拡張生成）に最適な Markdown チャンクを生成します。
- `--include-logic`: **[NEW]** エクセル内の計算式（例: `=SUM(A1:B1)`）と表示形式（円、%、日付）を抽出し、AI に伝えます。集計表の解析精度が劇的に向上します。
- `--no-vision`: 画像解析を無効化し、テキストと構造情報のみで実行します（高速・低コスト）。
- `--dpi`: 画像変換時の解像度を指定（デフォルト: 150）。細かい文字が多い場合は 300 を推奨。

## 3. 高度な抽出機能

本ツールは、単なるテキスト抽出ではなく、以下の 4 つの「インテリジェンス」を組み合わせて解析を行います。

### 3.1 視覚スタイル認識 (Style-aware)
セルの罫線の太さや背景色を CSS 形式で AI に伝えます。これにより、AI は「太い線で囲まれた範囲が一つのセクションである」ことを視覚的に理解します。

### 3.2 セマンティック・クリーニング (Semantic Cleaning)
「氏　　名」のように、レイアウト目的で挿入された不自然な空白を自動的に除去します。

### 3.3 Visual Intelligence (図表データ統合)
シート内の画像やグラフを VLM (Vision LLM) で解析し、その内容（数値データ等）を HTML テーブルの中に注釈として動的に埋め込みます。

### 3.4 ロジック・インジェクション (Logic-aware)
`--include-logic` フラグを有効にすると、セルの「計算式」を抽出します。
- AI は `=SUM(...)` を見ることで、そのセルが個別の値ではなく「合計値」であることを確信し、データの親子関係を正しく復元します。

## 4. プログラマ向けの使い方 (Python API)

ライブラリとして自身のプロジェクトに組み込むことも可能です。

```python
import asyncio
from kami_excel_extractor import KamiExcelExtractor, ExtractionOptions

async def main():
    extractor = KamiExcelExtractor(output_dir="output")
    options = ExtractionOptions(
        model="gemini-1.5-pro",
        include_logic=True,
        include_visual_summaries=True
    )
    
    result = await extractor.aextract_structured_data("report.xlsx", options=options)
    print(result["sheets"]["Sheet1"])

if __name__ == "__main__":
    asyncio.run(main())
```

## 5. 精度を最大化するためのコツ

1. **モデルの選択**: 複雑な神エクセルには `gemini-1.5-pro` などの、長いコンテキストと高い推論能力を持つモデルを推奨します。
2. **適切な DPI**: セル内の文字が非常に小さい場合は `--dpi 300` を指定してください。
3. **ロジックの活用**: 数値の整合性が重要な帳票では、必ず `--include-logic` を有効にしてください。
4. **システムプロンプトの調整**: 特殊な業務知識が必要な場合は、`--system-prompt "あなたは〇〇業界の専門家です..."` のように役割を指定してください。
