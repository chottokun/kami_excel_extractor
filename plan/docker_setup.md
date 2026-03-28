# **Dockerインフラ構成詳細**

本システムは Docker Compose を利用し、レンダリング環境（Pipeline Worker）と LLM ゲートウェイ（LiteLLM）を完全に分離して管理する。

## **1. docker-compose.yml 構成**

```yaml
version: '3.8'

services:
  litellm-proxy:
    image: ghcr.io/berriai/litellm:main
    container_name: litellm-proxy
    ports:
      - "4000:4000"
    volumes:
      - ./litellm-config.yaml:/app/config.yaml
    environment:
      - LLM_API_KEY=${LLM_API_KEY}
      - AZURE_API_KEY=${AZURE_API_KEY}
      - AZURE_API_BASE=${AZURE_API_BASE}
    command: ["--config", "/app/config.yaml", "--port", "4000"]
    restart: always

  pipeline-worker:
    build:
      context: .
      dockerfile: Dockerfile.worker
    container_name: pipeline-worker
    volumes:
      - ./src:/app/src
      - ./data/input:/app/input
      - ./data/output:/app/output
      - ./templates:/app/templates
    environment:
      - LITELLM_PROXY_URL=http://litellm-proxy:4000
      - PYTHONUNBUFFERED=1
    depends_on:
      - litellm-proxy
    restart: unless-stopped
```

## **2. Dockerfile.worker (Pipeline Worker用)**

LibreOffice と日本語フォントを含む堅牢な実行環境を構築する。

```dockerfile
FROM python:3.10-slim-bullseye

# システムパッケージとLibreOfficeのインストール
RUN apt-get update && apt-get install -y --no-install-recommends 
    libreoffice-calc 
    libreoffice-java-common 
    default-jre 
    fonts-noto-cjk 
    fonts-liberation 
    curl 
    && apt-get clean 
    && rm -rf /var/lib/apt/lists/*

# 作業ディレクトリの設定
WORKDIR /app

# 依存ライブラリのインストール
COPY pyproject.toml .
RUN pip install --no-cache-dir .

# ソースコードのコピー
COPY src/ /app/src/
COPY templates/ /app/templates/

# アプリケーションの実行
CMD ["python", "src/main.py"]
```

## **3. 設定のポイント**

### **A. フォントの重要性**
`fonts-noto-cjk` をインストールすることで、Excel 内の日本語文字が LibreOffice レンダリング時に文字化けしたり、レイアウトが崩れたりすることを防ぐ。

### **B. LiteLLM Proxy の利用**
Worker 側には個別の LLM SDK (Google / OpenAI) を入れず、標準の `openai` ライブラリのみを使用する。これにより、モデルの切り替えは `litellm-config.yaml` の変更のみで完結する。

### **C. ボリュームマウント**
- `input`: Excel ファイル投入用。
- `output`: 抽出結果 JSON およびデバッグ用 PNG 保存用。
- `templates`: ヘッダー・フッター無効設定済みの `.ots` (LibreOffice テンプレート) 格納用。
