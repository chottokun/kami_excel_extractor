# **Dockerインフラ構成詳細 (完全版)**

本システムは Docker Compose を利用し、LibreOffice を含む複雑な依存関係を完全にパッケージ化している。これにより、環境に依存せず高品質な Excel レンダリングと VLM 抽出を可能にする。

## **1. システム構成のポイント**

### **A. 多層変換エンジンの統合**
`Dockerfile.worker` 内に以下のツールを統合し、PDF 変換のフォールバックチェーンを 100% 機能させる。
- **LibreOffice (soffice)**: Excel から PDF への変換。
- **Poppler (pdftocairo)**: 高品質な PDF to PNG 変換（優先）。
- **Ghostscript & ImageMagick**: 最終的なフォールバック変換。セキュリティポリシーを調整し PDF 読み込みを許可済み。
- **fonts-noto-cjk**: 日本語 Excel の文字化けとレイアウト崩れを完全に防止。

### **B. uv による高速・堅牢なパッケージ管理**
 astral-sh/uv を採用し、ビルド時間の短縮と決定論的な依存関係（uv.lock）の解決を実現。

### **C. 権限の不整合防止 (UID/GID Mapping)**
ホストOSとコンテナ間でのファイル書き込み権限問題を避けるため、環境変数 `UID` / `GID` に基づいた実行ユーザーの動的生成を行う。

## **2. 使い方 (Docker Compose)**

### **A. 監視パイプラインの起動 (常駐型)**
`data/input` ディレクトリを監視し、ファイルが投入されるたびに自動解析を行う。

```bash
docker compose up -d --build
```

### **B. CLI による手動解析 (タスク型)**
特定のファイルに対して、個別のモデルやオプションを指定して即座に実行する。

```bash
# 基本実行 (data/input/sample.xlsx を解析)
docker compose run --rm cli sample.xlsx --model gemini/gemini-1.5-flash

# RAG チャンク生成を有効化
docker compose run --rm cli sample.xlsx --rag

# 画像解析を無効化 (テキストのみ)
docker compose run --rm cli sample.xlsx --no-vision
```

## **3. 運用設定 (.env)**

コンテナの挙動はプロジェクトルートの `.env` で制御する。

| 変数名 | 説明 | 推奨値 |
| :--- | :--- | :--- |
| `LLM_MODEL` | 使用する LLM モデル名 | `gemini/gemini-1.5-flash` |
| `LLM_API_KEY` | 各プロバイダーの API キー | (必須) |
| `LLM_RPM_LIMIT` | 1分あたりのリクエスト制限 | `1` (ローカルLLM時), `15` (API時) |
| `UID` / `GID` | ホストOSのユーザーID/グループID | `id -u` / `id -g` の結果 |

## **4. トラブルシューティング**

- **PDF 生成に失敗する場合**: コンテナ内の `soffice` が正常に起動しているか確認。`docker exec -it pipeline-worker-lib soffice --version` で確認可能。
- **画像抽出がスキップされる**: ログに `cannot identify image file` が出る場合、Excel 独自の図形形式（zlib圧縮）であり、現在の Pillow では非対応。構造化データ（JSON）自体は抽出されるため、運用上は問題ない。
