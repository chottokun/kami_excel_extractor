# RAG最適化ガイド (RAG Optimization Guide)

本ドキュメントは、Kami Excel Extractorが生成するRAG（検索拡張生成）用データ（Contextual Chunks）の仕様、アーキテクチャ、および主要なLLM/RAGフレームワーク（LangChain, LlamaIndex, Dify）との連携方法について解説します。

---

## 1. 背景と設計思想

従来のツールでExcelファイルをテキスト（Markdown等）に変換してRAGに投入した場合、以下の問題が発生していました：
- **結合セルやレイアウト構造の消失**: 表の親子関係や見出し構造が崩れ、LLMが誤った関連付けを行う。
- **メタデータの欠落**: チャンク分割されたテキスト単体では「元のシート名」「セル範囲」「抽出日時」などがわからず、検索時のフィルタリングやトレーサビリティが機能しない。
- **計算式・単位の欠落**: `=SUM(...)` などの計算式から得られる「この値は合計値である」という文脈や、通貨・パーセントといった単位が欠落し、数値の意味を誤認する。

本システムのRAG最適化モジュールは、これらの課題を解決するために**原本の文脈を保持したセグメント（Contextual Chunks）**を自動構成します。

---

## 2. Contextual Chunking 仕様

### 2.1 チャンク構造 (YAML Front Matter 形式)

`yaml_frontmatter` 形式では、各チャンクファイルの先頭にメタデータがYAML形式で埋め込まれ、下流のドキュメントローダーが自動的にメタデータフィールドとして解析できるように設計されています。

#### 出力例:
```markdown
---
chunk_index: 1
coordinates: B2:E15
extraction_date: '2026-06-27T22:04:52.810342'
has_formulas: true
has_media: true
section: 施工報告
sheet_name: 現場写真
source_file: complex_report.xlsx
total_chunks: 2
---

# 現場写真 > 施工報告

- **確認事項**: 外壁のクラック補修が完了。

> ℹ️ セル E15 は計算式 `=SUM(E3:E14)`（単位: JPY）から導出された集計値です。
```

### 2.2 メタデータフィールドの定義

| フィールド名 | 型 | 説明 |
|:---|:---|:---|
| `source_file` | string | 抽出元のExcelファイル名。 |
| `sheet_name` | string | 抽出元のワークシート名。 |
| `section` | string | チャンクが属するMarkdownの論理セクション（見出し）。 |
| `coordinates` | string | チャンク内のテキストに対応するExcel上のセル座標範囲（例: `A1:F12`）。 |
| `has_formulas` | boolean | チャンクのセル範囲に何らかの計算式が含まれているか。 |
| `has_media` | boolean | チャンクのセル範囲内、またはテキスト中に関連する画像/図表が含まれているか。 |
| `extraction_date` | string (ISO8601) | パイプラインが抽出を実行した日時。 |
| `chunk_index` | integer | 該当シート内におけるこのチャンクの連番（1-indexed）。 |
| `total_chunks` | integer | 該当シートにおける総チャンク数。 |

---

## 3. インライン論理注釈 (Logic Annotation)

`--include-logic` が有効である場合、`ContextualChunkGenerator` はExcel内の主要な集計式（`SUM`, `AVERAGE`, `COUNT`, `MAX`, `MIN`, `SUBTOTAL`, `VLOOKUP`, `IF` 等）を自動判定します。

判定された計算式は、ハルシネーションを抑制するための文脈情報として、**Markdownチャンクの末尾に引用注釈として動的にインジェクション**されます。
```markdown
> ℹ️ セル {座標} は計算式 `{Excel計算式}`{単位情報} から導出された集計値です。
```
これによって、LLMは「この値が個別データではなく、全体の合計や判定ロジックを反映したものである」というメタ文脈を正しく理解できます。

---

## 4. JSONL 出力スキーマ

`--rag-format jsonl` を指定した場合、すべてのチャンクは1行ごとにシリアライズされたJSONオブジェクトとして `{stem}_rag.jsonl` に一括保存されます。

```json
{
  "content": "# 現場写真 > 施工報告\n- **確認事項**: 外壁のクラック補修が完了。\n...",
  "metadata": {
    "source_file": "complex_report.xlsx",
    "sheet_name": "現場写真",
    "section": "施工報告",
    "coordinates": "B2:E15",
    "has_formulas": true,
    "has_media": true,
    "extraction_date": "2026-06-27T22:04:52.810342",
    "chunk_index": 1,
    "total_chunks": 2
  }
}
```

---

## 5. 主要RAGフレームワークとの連携方法

### 5.1 LangChain との連携 (Python)

LangChainでは、YAML Front Matterを自動パースして `Document` オブジェクトの `metadata` に格納することができます。

```python
from pathlib import Path
from langchain_community.document_loaders import TextLoader
from langchain_core.documents import Document
import yaml

def load_contextual_chunk(file_path: Path) -> Document:
    """YAML Front Matter付きのMarkdownチャンクを読み込む"""
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Front Matter のパース
    if content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) >= 3:
            metadata = yaml.safe_load(parts[1])
            body = parts[2].strip()
            return Document(page_content=body, metadata=metadata)
            
    return Document(page_content=content, metadata={})

# ディレクトリ以下の全チャンクをロード
rag_dir = Path("output/complex_report_rag")
documents = [load_contextual_chunk(p) for p in rag_dir.glob("*.md")]

# ベクトルストアへ追加
# db = Chroma.from_documents(documents, embeddings)
```

### 5.2 LlamaIndex との連携 (Python)

LlamaIndexでは、カスタム `Node` メタデータパーサーを介して、Front Matterから抽出したメタデータを構造化フィルタリングに使用できます。

```python
import yaml
from pathlib import Path
from llama_index.core import Document
from llama_index.core.schema import MetadataMode

def load_llama_documents(rag_dir_path: str):
    documents = []
    for p in Path(rag_dir_path).glob("*.md"):
        with open(p, "r", encoding="utf-8") as f:
            raw_text = f.read()
            
        if raw_text.startswith("---"):
            parts = raw_text.split("---", 2)
            if len(parts) >= 3:
                metadata = yaml.safe_load(parts[1])
                body = parts[2].strip()
                
                doc = Document(
                    text=body,
                    metadata=metadata,
                    # LLMに渡す際、または埋め込み時に除外するメタデータの設定
                    excluded_embed_metadata_keys=["extraction_date", "total_chunks"],
                    excluded_llm_metadata_keys=["extraction_date"]
                )
                documents.append(doc)
    return documents

# ロードとインデックス化
docs = load_llama_documents("output/complex_report_rag")
# index = VectorStoreIndex.from_documents(docs)

# メタデータによるフィルタリングクエリの例（特定のシートのみを検索対象にする）
# from llama_index.core.vector_stores import MetadataFilters, ExactMatchFilter
# filters = MetadataFilters(filters=[ExactMatchFilter(key="sheet_name", value="現場写真")])
# retriever = index.as_retriever(filters=filters)
```

### 5.3 Dify との連携 (ノーコード/ローコード)

Difyでは、生成された **JSONLファイル** または **Markdownファイル群** を「ナレッジ (Knowledge)」として取り込むことで、Excelのセマンティクスをそのまま検索に使用できます。

#### 連携手順:

1. **JSONLによる一括アップロード (推奨)**:
   - Difyの管理画面から **「ナレッジ」-> 「ナレッジを作成」-> 「テキストファイルからインポート」** を選択します。
   - 本システムで出力した `{stem}_rag.jsonl` をアップロードします。
   - アップロード設定画面にて、「**JSONL構造の解析**」を選択します。
   - `content` フィールドを **本文** に、`metadata` フィールドを **メタデータキー** としてマッピングします。
   
2. **メタデータフィルタの適用**:
   - Difyのチャットフロー（Workflow）やエージェントの設定内で、作成したナレッジを紐付けます。
   - ナレッジ検索ノードの「**フィルター条件**」に `sheet_name == "現場写真"` や `has_formulas == true` を追加して、検索対象を自動絞り込みできます。
