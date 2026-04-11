# **システムアーキテクチャ定義 (実証済み)**

本システムは、Excelドキュメントの視覚情報（画像）と論理構造（メタデータ・埋め込みメディア）を組み合わせ、VLM（Vision Language Model）を用いて高度な構造化データ抽出を実現する。

## **1. システム全体図**

```mermaid
graph TD
    User([Excel File]) --> Worker[Pipeline Worker Container]
    
    subgraph "kami_excel_extractor Library"
        Worker --> |1. Local Render| LO[ExcelConverter]
        Worker --> |2. Structural Analysis| OPX[MetadataExtractor]
        Worker --> |3. Multi-modal Request| LiteLLM[LiteLLM Library Integration]
        
        LO --> |PDF -> PNG Image| Worker
        OPX --> |HTML Table + Extracted Media| Worker
    end
    
    LiteLLM --> |4. Standardized API Call| VLM[VLM: Gemini / OpenAI / Azure]
    VLM --> |5. Structured YAML| LiteLLM
    LiteLLM --> |6. Final Parsed JSON| Worker
    Worker --> Output([Structured JSON/YAML Data])
```

## **2. 実装済みコンポーネントの役割**

### **A. ExcelConverter (LibreOffice + Multi-Engine Fallback)**
- **役割:** Excelを物理画像に変換。
- **2-Step Process:** ヘッドレス環境でのCalcの不安定さを回避するため、「Excel -> PDF (soffice)」→「PDF -> PNG」のフローを採用。
- **3-Layer Fallback Chain:** PDFからPNGへの変換において、以下のエンジンを順次試行し、環境依存の失敗を最小化。
    1.  **pdftocairo (Poppler)**: 高品質なベクターレンダリング。
    2.  **PyMuPDF (fitz)**: Pythonネイティブな高速レンダリング（pdftocairo不在時のバックアップ）。
    3.  **ImageMagick (magick/convert)**: 汎用的な最終フォールバック。
- **並行処理:** UUIDベースの一時プロファイル（UserInstallation）による隔離実行。

### **B. MetadataExtractor (openpyxl + Media Logic)**
- **論理抽出:** セル結合やスタイルを維持したまま、巨大なシートも軽量な **HTMLの `<table>` タグ** として構造化抽出（Promptのトークン量を約90%削減）。
- **メディア抽出:** シート内に埋め込まれた画像（現場写真等）を物理ファイルとして抽出し、紐付け。Pillowによるフォーマット正規化を実施。

### **C. KamiExcelExtractor (Core Orchestrator)**
- **構成管理:** `ExtractionOptions` および `RagOptions` (Pydantic) による柔軟なパラメータ制御。
- **統合ロジック:** 視覚情報と論理情報を統合し、マルチモーダル・プロンプト（OpenAI互換形式）を構成。
- **モデル自動判別:** 利用可能な最新のFlashモデルを自動的に選択。
- **非同期オーケストレーション:**
    - **並列解析:** `asyncio.gather` を用いて、複数シートの解析や画像のビジュアルサマリー生成を並列に実行。
    - **レート制限:** `LLM_RPM_LIMIT` 環境変数に基づいた `asyncio.Semaphore` により、外部APIへの過剰なリクエストを制御。
    - **スレッドオフロード:** 重い画像エンコード（Base64）やファイル I/O は `asyncio.to_thread` を介してスレッドプールで実行し、メインのスレッドをブロックしない。
    - **推論バイパス:** `is_simple` フラグが立ったシート（単純な表構造）は LLM 呼び出しをスキップし、直接抽出データを使用。

### **D. LiteLLM Integration**
- **役割:** プロバイダー（Google, OpenAI, Microsoft）のAPI差異をライブラリ層で吸収。
- **柔軟性:** `litellm.acompletion` を使用した非同期呼び出しに対応。モデル名のプレフィックスを変えるだけでモデル切替が可能。

## **3. 実証されたデータフロー**

1.  **入力:** Excelファイル（方眼紙、複数シート、写真入り等）を検知。
2.  **前処理:** 
    - PDFを経由して高品質なPNGを生成。
    - XMLをパースして全シートの「HTMLテーブル（情報の地図）」と「写真ファイル」を抽出。
3.  **VLM推論:** LiteLLMを介して、画像とHTMLを統合したプロンプトを送信。出力形式には軽量でトークン溢れを防ぐ**YAML**を指定。
4.  **出力:** 文脈を理解した精緻な構造化YAMLを出力し、プログラム側でJSONや専用ディレクトリへ安全に保存。
