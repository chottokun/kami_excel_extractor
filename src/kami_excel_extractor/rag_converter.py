import json
from typing import Any, Dict, List, Optional
from pathlib import Path

class JsonToMarkdownConverter:
    """JSONをRAGに適したMarkdown形式に変換するクラス"""

    def convert(self, data: Any, level: int = 1) -> str:
        if isinstance(data, dict):
            return self._convert_dict(data, level)
        elif isinstance(data, list):
            return self._convert_list(data, level)
        else:
            return str(data)

    def _convert_dict(self, data: Dict[str, Any], level: int) -> str:
        lines = []
        header = "#" * min(level + 1, 6)
        for key, value in data.items():
            if key == "media" and isinstance(value, list):
                lines.append(f"\n{header} 関連メディア\n")
                lines.append(self._convert_media(value))
            elif isinstance(value, (dict, list)):
                lines.append(f"\n{header} {key}\n")
                lines.append(self.convert(value, level + 1))
            else:
                lines.append(f"- **{key}**: {value}")
        return "\n".join(lines)

    def _convert_list(self, data: List[Any], level: int) -> str:
        if not data:
            return ""
        
        # リストの要素が辞書で、かつ同じキーを持っている場合はテーブル化を試みる
        if all(isinstance(item, dict) for item in data) and len(data) > 0:
            keys = data[0].keys()
            if all(item.keys() == keys for item in data):
                return self._convert_to_table(data, keys)
        
        lines = []
        for item in data:
            if isinstance(item, (dict, list)):
                lines.append(f"\n---\n")
                lines.append(self.convert(item, level))
            else:
                lines.append(f"- {item}")
        return "\n".join(lines)

    def _convert_to_table(self, data: List[Dict[str, Any]], keys: Any) -> str:
        header_line = "| " + " | ".join(keys) + " |"
        separator_line = "| " + " | ".join(["---"] * len(keys)) + " |"
        rows = []
        for item in data:
            row = "| " + " | ".join(str(item.get(k, "")) for k in keys) + " |"
            rows.append(row)
        return "\n".join([header_line, separator_line] + rows)

    def _convert_media(self, media_list: List[Dict[str, Any]]) -> str:
        lines = []
        for item in media_list:
            filename = item.get("filename", "")
            summary = item.get("visual_summary", "").strip()
            
            # Markdown image syntax and summary
            # Note: Using relative path 'media/' as it will be resolved relative to the output md/pdf
            lines.append(f"![画像](media/{filename})")
            if summary:
                # Ensure the summary starts with [画像概要] if it wasn't already there (though the prompt asks for it)
                if not summary.startswith("[画像概要]"):
                    summary = f"**[画像概要]**: {summary}"
                else:
                    # Bold the prefix and ensure colon for consistency
                    summary = summary.replace("[画像概要]", "**[画像概要]**:", 1)
                lines.append(f"{summary}\n")
        return "\n".join(lines)

class RagChunker:
    """Markdownを見出しベースでチャンク分割するクラス"""

    def __init__(self, metadata: Optional[Dict[str, Any]] = None):
        self.metadata = metadata or {}

    def chunk(self, markdown_text: str, source_name: str) -> List[Dict[str, Any]]:
        chunks = []
        lines = markdown_text.split("\n")
        current_chunk_lines = []
        current_header = ""
        
        for line in lines:
            if line.startswith("# "): # Top level header
                if current_chunk_lines:
                    chunks.append(self._create_chunk(current_chunk_lines, current_header, source_name))
                current_chunk_lines = [line]
                current_header = line.lstrip("# ").strip()
            else:
                current_chunk_lines.append(line)
        
        if current_chunk_lines:
            chunks.append(self._create_chunk(current_chunk_lines, current_header, source_name))
            
        return chunks

    def _create_chunk(self, lines: List[str], header: str, source_name: str) -> Dict[str, Any]:
        content = "\n".join(lines).strip()
        chunk_metadata = self.metadata.copy()
        chunk_metadata.update({
            "source": source_name,
            "section": header
        })
        return {
            "content": content,
            "metadata": chunk_metadata
        }
