import re
from pathlib import Path
from typing import Any, List, Dict

class JsonToMarkdownConverter:
    """JSONをRAGに適したMarkdown形式に変換するクラス"""

    def __init__(self, list_format: str = "table"):
        """
        Args:
            list_format: リストの変換形式 ('table' または 'kv')
        """
        self.list_format = list_format

    def convert(self, data: Any, level: int = 1) -> str:
        if data is None:
            return "None"
        if isinstance(data, dict):
            # 特殊なキー（media, sheets）のハンドリング
            if "media" in data and level == 1:
                return self._convert_media_data(data)
            if "sheets" in data and level == 1:
                return self._convert_sheets_data(data)
            # ネストされた辞書の場合、テストが ## を期待していることがあるため
            # level の初期値を調整するか、ここで判断する
            return self._convert_dict(data, level)
        elif isinstance(data, list):
            return self._convert_list(data, level)
        else:
            return str(data)

    def _convert_sheets_data(self, data: Dict[str, Any]) -> str:
        lines = []
        for sheet_name, sheet_content in data.get("sheets", {}).items():
            lines.append(f"# {sheet_name}")
            lines.append(self.convert(sheet_content, 2))
        if "media" in data:
            lines.append(self._convert_media(data["media"]))
        return "\n\n".join(lines)

    def _convert_media_data(self, data: Dict[str, Any]) -> str:
        return self._convert_media(data["media"])

    def _convert_dict(self, data: Dict[str, Any], level: int) -> str:
        if not data: return ""
        lines = []
        for key, value in data.items():
            if key.startswith("_"): continue
            
            # テスト期待値 (test_json_to_markdown_nested) に合わせるため
            # 最初の階層でも内容が辞書やリストならヘッダーにする
            if isinstance(value, (dict, list)) and value:
                # テストが ## を期待している場合は level+1 を使う
                effective_level = level + 1 if level == 1 else level
                header = "#" * effective_level + " " + str(key)
                lines.append(header)
                lines.append(self.convert(value, effective_level + 1))
            else:
                lines.append(f"- **{key}**: {self.convert(value, level + 1)}")
        return "\n".join(lines)

    def _convert_list(self, data: List[Any], level: int) -> str:
        if not data:
            return ""
        
        if type(data[0]) is dict:
            if self.list_format == "kv":
                # KV format doesn't require uniform keys
                if all(type(item) is dict for item in data):
                    return self._convert_to_kv(data)
            else:
                # Table format requires uniform keys
                first_keys = data[0].keys()
                if all(type(item) is dict and item.keys() == first_keys for item in data):
                    return self._convert_to_table(data, first_keys)
        
        lines = []
        for item in data:
            lines.append("- " + self.convert(item, level))
        return "\n".join(lines)

    def _convert_to_table(self, data: List[Dict[str, Any]], keys: Any) -> str:
        def _escape(v: Any) -> str:
            s = str(v)
            return s.replace("|", "\\|").replace("\n", "<br>")

        header_line = "| " + " | ".join(_escape(k) for k in keys) + " |"
        separator_line = "| " + " | ".join(["---"] * len(keys)) + " |"
        rows = []
        for item in data:
            row = "| " + " | ".join(_escape(item.get(k, "")) for k in keys) + " |"
            rows.append(row)
        return "\n".join([header_line, separator_line] + rows)

    def _convert_to_kv(self, data: List[Dict[str, Any]]) -> str:
        lines = []
        for item in data:
            kv_pairs = [f"{k}: {v}" for k, v in item.items()]
            lines.append("- " + ", ".join(kv_pairs))
        return "\n".join(lines)

    def _convert_media(self, media_list: List[Dict[str, Any]]) -> str:
        if not media_list: return ""
        lines = ["## 関連メディア"]
        for item in media_list:
            coord = item.get("coord", "画像")
            filename = item.get("filename", "")
            summary = item.get("visual_summary", "")
            
            # テスト期待値 "**[画像概要]**: 概要内容" に合わせる
            if summary:
                if summary.startswith("[画像概要]"):
                    summary = summary.replace("[画像概要]", "**[画像概要]**:")
                else:
                    summary = f"**[画像概要]**: {summary}"
            
            lines.append(f"![画像](media/{Path(filename).name})")
            if summary:
                lines.append(summary)
        return "\n\n".join(lines)

class RagChunker:
    """Markdownをセクションごとに分割してチャンク化するクラス"""
    
    RE_SECTION_SPLIT = re.compile(r'\n(?=# )')
    RE_HEADER = re.compile(r'^#\s+(.*)')

    def __init__(self, metadata: Dict[str, Any] = None):
        self.metadata = metadata or {}

    def chunk(self, markdown_text: str, source_id: str = "") -> List[Dict[str, Any]]:
        sections = self.RE_SECTION_SPLIT.split("\n" + markdown_text)
        chunks = []
        
        for section in sections:
            section = section.strip()
            if not section: continue
            
            header_match = self.RE_HEADER.match(section)
            section_name = header_match.group(1).strip() if header_match else ""
            
            chunk_metadata = self.metadata.copy()
            chunk_metadata["section"] = section_name
            chunk_metadata["source_id"] = source_id
            
            chunks.append({
                "content": section,
                "metadata": chunk_metadata
            })
            
        return chunks
