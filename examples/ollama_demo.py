import logging
import json
import asyncio
from pathlib import Path
from kami_excel_extractor.core import KamiExcelExtractor
from kami_excel_extractor.schema import ExtractionOptions, RagOptions

# ロギング設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

async def run_ollama_demo():
    # 1. 初期化 (サーバー上の Ollama エンドポイントと Qwen 3.5:4b を指定)
    # ここではローカルの Ollama サーバーを利用します
    extractor = KamiExcelExtractor(
        base_url="http://localhost:11434",
        output_dir="data/output/ollama_demo"
    )

    # モデルを指定 (ollama/ プレフィックスを付けることで LiteLLM が Ollama 用のリクエストを生成)
    target_model = "ollama/qwen3.5:4b"
    
    # 2. 対象の Excel ファイル
    excel_path = "sample_hoganshi.xlsx"
    if not Path(excel_path).exists():
        logger.error(f"Excel file not found: {excel_path}")
        return

    logger.info(f"Starting extraction using {target_model}...")

    # 3. 構造化データの抽出 (テキスト抽出のみを確認するため include_visual_summaries=False)
    # 実際には Excel の HTML 表現が LLM に送信されます
    try:
        extract_options = ExtractionOptions(
            model=target_model,
            include_visual_summaries=False
        )
        results = await extractor.aextract_structured_data(
            excel_path,
            options=extract_options
        )

        # 4. 結果の表示
        logger.info("Extraction completed results:")
        print(json.dumps(results, indent=2, ensure_ascii=False))
        
        # 5. RAG チャンクの生成デモ
        logger.info("Generating RAG chunks for downstream tasks...")
        rag_options = RagOptions(model=target_model)
        chunks_map, structured_data = await extractor.aextract_rag_chunks(
            excel_path,
            options=rag_options
        )
        
        for sheet_name, data in chunks_map.items():
            print(f"\n--- Sheet: {sheet_name} ---")
            print(f"Number of chunks: {len(data['chunks'])}")
            if data['chunks']:
                print(f"Sample chunk: {data['chunks'][0].page_content[:200]}...")

    except Exception as e:
        logger.error(f"An error occurred during extraction: {e}")

if __name__ == "__main__":
    asyncio.run(run_ollama_demo())
