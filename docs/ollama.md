# Ollama (ローカル LLM) の利用方法

本ツールは LiteLLM を介して Ollama に対応しており、ローカル環境で動作する LLM を使用して Excel の構造化抽出を行うことが可能です。

## 1. セットアップ

### Ollama のインストール
[Ollama 公式サイト](https://ollama.com/) から OS に合わせたインストーラーをダウンロードし、インストールしてください。

### モデルのプル
本システムで動作確認済みのモデルをプルします。

```bash
# テキスト抽出用 (推奨: Qwen シリーズ)
ollama pull qwen3.5:4b

# 画像解析 (Vision) 用 (必要な場合)
ollama pull llava
```

## 2. 実行設定

### 基本的な使い方
`KamiExcelExtractor` の初期化時に `base_url` を指定し、抽出実行時に `ollama/` プレフィックスを付けたモデル名を指定します。

```python
from kami_excel_extractor import KamiExcelExtractor

# Ollama サーバーの URL を指定 (デフォルトは http://localhost:11434)
extractor = KamiExcelExtractor(base_url="http://localhost:11434")

# テキスト抽出を実行 (qwen3.5:4b 等)
# include_visual_summaries=False にすることでテキストのみの抽出が可能です
result = extractor.extract_structured_data(
    "path/to/excel.xlsx",
    model="ollama/qwen3.5:4b",
    include_visual_summaries=False
)
print(result)
```

## 3. Vision (VLM) 機能の使用方法

グラフや画像などの視覚情報を解析する場合、以下の手順が必要です。

1.  **Vision 対応モデルの指定**: `model="ollama/llava"` など、画像認識に対応したモデルを指定してください。
2.  **フラグの有効化**: `include_visual_summaries=True` をセットしてください。

```python
```python
# Vision モデルを使用する場合の例
result = extractor.extract_structured_data(
    "path/to/excel.xlsx",
    model="ollama/llava", # Vision 対応モデル
    include_visual_summaries=True
)
```

## 4. CLI での利用 (推奨)

CLI を使用すると、コマンドラインから直感的に Ollama を呼び出せます。

```bash
# qwen3.5:4b を使用してテキストのみで抽出
kami-excel sample.xlsx --model ollama/qwen3.5:4b --no-vision --base-url http://localhost:11434
```

## 5. 推奨モデル

- **テキスト抽出 (YAML 生成)**:
    - `qwen3.5:4b`: 高速かつ指示追従性が非常に高い。
    - `qwen2.5-coder:7b`: より複雑な構造の解析に適している。
- **画像解析 (Vision)**:
    - `llava`: 一般的な画像説明に適している。
    - `moondream`: 軽量かつ高性能な Vision モデル。

## 💡 ヒント

- **混在利用**: データの抽出はローカルの Ollama で行い、高難易度な画像解析だけ Gemini を使用するといったハイブリッドな運用も可能です。
- **パフォーマンス**: 4B 程度のモデルであれば、ミドルスペックの PC でも高速に動作します。
