import json
from pathlib import Path
from typing import Any, Dict, List, Union


class JsonlExporter:
    """RAGチャンクをJSON Lines (JSONL) 形式でエクスポートするクラス"""

    @staticmethod
    def export(chunks: List[Dict[str, Any]], output_path: Union[str, Path]) -> None:
        """チャンクのリストをJSONLファイルとして保存する"""
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, "w", encoding="utf-8") as f:
            for chunk in chunks:
                line_data = {
                    "content": chunk.get("content", ""),
                    "metadata": chunk.get("metadata", {})
                }
                f.write(json.dumps(line_data, ensure_ascii=False) + "\n")

    @staticmethod
    def to_jsonl_string(chunks: List[Dict[str, Any]]) -> str:
        """チャンクのリストをJSONL形式の文字列に変換する"""
        lines = []
        for chunk in chunks:
            line_data = {
                "content": chunk.get("content", ""),
                "metadata": chunk.get("metadata", {})
            }
            lines.append(json.dumps(line_data, ensure_ascii=False))
        return "\n".join(lines) + "\n"
