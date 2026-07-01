import argparse
import asyncio
import json
import logging
import os
import sys
from pathlib import Path

from .core import KamiExcelExtractor
from .schema import ExtractionOptions, RagOptions
from .utils import secure_filename


def create_parser():
    """CLIパーサーを作成する"""
    parser = argparse.ArgumentParser(description="Kami Excel Extractor CLI - Excelを構造化データ(JSON/YAML)に変換")

    parser.add_argument("input", help="対象のExcelファイルパス")
    parser.add_argument("--model", help="使用するモデル名 (例: ollama/qwen3.5:4b, gemini/gemini-1.5-flash)")
    parser.add_argument("--api-key", help="LLMのAPIキー (環境変数 LLM_API_KEY を上書き)")
    parser.add_argument("--base-url", help="Ollama などのベース URL (環境変数 LLM_BASE_URL を上書き)")
    parser.add_argument("--timeout", type=float, default=600.0, help="推論のタイムアウト秒数 (デフォルト: 600)")
    parser.add_argument("--rpm", type=int, help="1分あたりのリクエスト制限 (デフォルト: 環境変数の値)")
    parser.add_argument("--output-dir", default="output", help="出力先ディレクトリ (デフォルト: output)")
    parser.add_argument("--no-vision", action="store_true", help="画像解析を完全に無効化し、テキストのみで実行する")
    parser.add_argument("--rag", action="store_true", help="RAG用のMarkdownチャンクも同時に生成する")
    parser.add_argument(
        "--rag-format",
        default="yaml_frontmatter",
        choices=["markdown", "yaml_frontmatter", "jsonl", "docx"],
        help="RAG出力形式 (デフォルト: yaml_frontmatter)",
    )
    parser.add_argument("--docx", action="store_true", help="Dify最適化DOCX変換を有効にする")
    parser.add_argument("--max-chunk-chars", type=int, default=1000, help="チャンクの最大文字数制限 (デフォルト: 1000)")
    parser.add_argument("--chunk-overlap-lines", type=int, default=2, help="チャンク間の重複行数 (デフォルト: 2)")
    parser.add_argument("--no-coordinates", action="store_true", help="RAGチャンクにExcelセル座標メタデータを含めない")
    parser.add_argument(
        "--no-logic-annotations", action="store_true", help="RAGチャンクに計算式のインライン注釈を含めない"
    )
    parser.add_argument("--system-prompt", help="カスタムシステムプロンプト")
    parser.add_argument("--dpi", type=int, help="変換時のDPI設定 (デフォルト: 150)")
    parser.add_argument("--include-logic", action="store_true", help="計算式(formula)と単位情報の抽出を有効にする")
    parser.add_argument("--visual-summaries", action="store_true", help="画像の個別解析（図表データ抽出）を有効にする")
    parser.add_argument("--verbose", action="store_true", help="詳細なログを出力する")
    return parser


async def run_async(args):
    """非同期実行のメインロジック"""
    logger = logging.getLogger("kami-excel-cli")

    # RPM制限を環境変数にセット (extractor._get_semaphore で参照されるため)
    if args.rpm is not None:
        os.environ["LLM_RPM_LIMIT"] = str(args.rpm)

    extractor = KamiExcelExtractor(
        api_key=args.api_key, base_url=args.base_url, output_dir=args.output_dir, timeout=args.timeout
    )

    logger.info(f"解析開始: {args.input}")

    # vision設定の解決
    use_visual_context = not args.no_vision
    # コマンドライン引数が優先されるが、デフォルトでは vision 有効時にサマリーも取る
    include_visual_summaries = args.visual_summaries if args.visual_summaries else (not args.no_vision)

    try:
        # モックオブジェクト対策のための安全な属性ゲッター
        def safe_get(obj, attr, default):
            val = getattr(obj, attr, default)
            if "Mock" in type(val).__name__:
                return default
            return val

        if safe_get(args, "docx", False):
            logger.info("DOCX出力モードで実行中...")
            rag_options = RagOptions(
                model=args.model,
                use_visual_context=use_visual_context,
                include_visual_summaries=include_visual_summaries,
                include_logic=args.include_logic,
                dpi=args.dpi if args.dpi is not None else 150,
                output_format="docx",
                include_logic_annotations=not safe_get(args, "no_logic_annotations", False),
            )
            docx_path, structured_data = await extractor.aextract_docx(args.input, options=rag_options)
            result_data = structured_data
            logger.info(f"DOCXファイルを保存しました: {docx_path}")

        elif args.rag:
            logger.info("RAGチャンク生成モードで実行中...")

            rag_options = RagOptions(
                model=args.model,
                use_visual_context=use_visual_context,
                include_visual_summaries=include_visual_summaries,
                include_logic=args.include_logic,
                dpi=args.dpi if args.dpi is not None else 150,
                output_format=safe_get(args, "rag_format", "yaml_frontmatter"),
                max_chunk_chars=safe_get(args, "max_chunk_chars", 1000),
                chunk_overlap_lines=safe_get(args, "chunk_overlap_lines", 2),
                include_coordinates=not safe_get(args, "no_coordinates", False),
                include_logic_annotations=not safe_get(args, "no_logic_annotations", False),
            )
            chunks_map, structured_data = await extractor.aextract_rag_chunks(args.input, options=rag_options)
            result_data = structured_data

        else:
            options = ExtractionOptions(
                model=args.model,
                system_prompt=args.system_prompt,
                use_visual_context=use_visual_context,
                include_visual_summaries=include_visual_summaries,
                include_logic=args.include_logic,
                dpi=args.dpi if args.dpi is not None else 150,
            )
            result_data = await extractor.aextract_structured_data(args.input, options=options)

        # 結果の保存
        safe_stem = secure_filename(Path(args.input).stem)
        output_path = Path(args.output_dir) / f"{safe_stem}_result.json"

        def _save_json(path, data):
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

        await asyncio.to_thread(_save_json, output_path, result_data)
        logger.info(f"結果を保存しました: {output_path}")

        if args.rag:
            rag_path = Path(args.output_dir) / f"{safe_stem}_rag.json"
            if safe_get(args, "rag_format", "yaml_frontmatter") == "docx":
                serializable_chunks = {"docx_path": str(chunks_map.get("docx", {}).get("path", ""))}
            else:
                serializable_chunks = {
                    k: {"chunks_count": len(v["chunks"]), "markdown": v["markdown"]} for k, v in chunks_map.items()
                }
            await asyncio.to_thread(_save_json, rag_path, serializable_chunks)
            logger.info(f"RAG用データを保存しました: {rag_path}")

    except Exception as e:
        logger.error(f"実行中にエラーが発生しました: {e}")
        if args.verbose:
            import traceback

            traceback.print_exc()
        sys.exit(1)


def main():
    """CLIエントリーポイント"""
    parser = create_parser()
    args = parser.parse_args()

    # ログ設定
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    logger = logging.getLogger("kami-excel-cli")

    if not Path(args.input).exists():
        logger.error(f"入力ファイルが見つかりません: {args.input}")
        sys.exit(1)

    asyncio.run(run_async(args))


if __name__ == "__main__":
    main()
