import subprocess
import tempfile
from pathlib import Path
import re
import logging
import shutil
import html

logger = logging.getLogger(__name__)

class DocumentGenerator:
    """MarkdownからHTMLを経由してPDFを生成するクラス"""

    # PR #9: Pre-compiled regexes
    RE_TABLE_SEP = re.compile(r'^:?-{2,}:?$')
    RE_HEADER = re.compile(r'^#+')
    RE_BOLD = re.compile(r'\*\*(.*?)\*\*')
    RE_IMAGE = re.compile(r'!\[.*?\]\((.*?)\)')

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)

    def _simple_md_to_html(self, md_text: str) -> str:
        """Markdown要素をHTMLに変換する (テーブル・画像対応強化)"""
        lines = md_text.splitlines()
        html_output = []
        in_table = False
        table_rows = []
        in_list = False

        for line in lines:
            stripped = line.strip()
            
            # テーブルの開始/継続判定
            if stripped.startswith("|") and stripped.endswith("|"):
                in_list = self._close_list_if_needed(html_output, in_list)
                if not in_table:
                    in_table = True
                    table_rows = []
                
                # 区切り行判定
                cells = [c.strip() for c in stripped.split("|")[1:-1]]
                if all(self.RE_TABLE_SEP.match(cell) for cell in cells):
                    continue
                table_rows.append(stripped)
                continue
            else:
                if in_table:
                    html_output.append(self._render_table(table_rows))
                    in_table = False
                    table_rows = []
            
            # 空行
            if not stripped:
                in_list = self._close_list_if_needed(html_output, in_list)
                continue
                
            # Headers
            if stripped.startswith("#"):
                in_list = self._close_list_if_needed(html_output, in_list)
                html_output.append(self._render_header(stripped))
            # Lists
            elif stripped.startswith("- "):
                if not in_list:
                    html_output.append("<ul>")
                    in_list = True
                html_output.append(self._render_list_item(stripped))
            # Images
            elif self.RE_IMAGE.match(stripped):
                in_list = self._close_list_if_needed(html_output, in_list)
                html_output.append(self._render_image_element(stripped))
            # Text
            else:
                in_list = self._close_list_if_needed(html_output, in_list)
                html_output.append(self._render_paragraph(stripped))

        # 末尾処理
        if in_table:
            html_output.append(self._render_table(table_rows))
        self._close_list_if_needed(html_output, in_list)

        return self._get_html_template("\n".join(html_output))

    def _close_list_if_needed(self, html_output: list, in_list: bool) -> bool:
        """リストが開始されている場合、閉じタグを追加する"""
        if in_list:
            html_output.append("</ul>")
        return False

    def _render_header(self, stripped_line: str) -> str:
        """ヘッダー要素をレンダリングする"""
        header_match = self.RE_HEADER.match(stripped_line)
        level = len(header_match.group()) if header_match else 1
        content = stripped_line.lstrip("#").strip()
        # 🔒 Security Fix: HTML escape before applying inline styles
        escaped_content = html.escape(content)
        return f"<h{level}>{self._apply_inline_styles(escaped_content)}</h{level}>"

    def _render_list_item(self, stripped_line: str) -> str:
        """リストアイテムをレンダリングする"""
        content = stripped_line[2:].strip()
        # 🔒 Security Fix: HTML escape before applying inline styles
        escaped_content = html.escape(content)
        return f"<li>{self._apply_inline_styles(escaped_content)}</li>"

    def _render_image_element(self, stripped_line: str) -> str:
        """画像要素をレンダリングする"""
        img_match = self.RE_IMAGE.search(stripped_line)
        img_path = img_match.group(1)
        # 🔒 Security Fix: HTML escape image source attribute
        escaped_img_path = html.escape(img_path, quote=True)
        return f'<div class="image-container"><img src="{escaped_img_path}" alt="画像"></div>'

    def _render_paragraph(self, stripped_line: str) -> str:
        """段落要素をレンダリングする"""
        # 🔒 Security Fix: HTML escape before applying inline styles
        escaped_text = html.escape(stripped_line)
        return f"<p>{self._apply_inline_styles(escaped_text)}</p>"

    def _get_html_template(self, body_html: str) -> str:
        """HTMLテンプレートを生成する"""
        return f"""<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: "Noto Sans CJK JP", "Noto Sans JP", "IPAGothic", sans-serif;
            line-height: 1.8;
            padding: 40px;
            color: #333;
            font-size: 14px;
        }}
        h1 {{ border-bottom: 3px solid #336699; padding-bottom: 10px; color: #336699; margin-top: 40px; font-size: 24px; }}
        h2 {{ border-left: 5px solid #336699; padding-left: 12px; margin-top: 30px; color: #336699; font-size: 20px; }}
        h3 {{ color: #444; margin-top: 25px; font-size: 16px; }}
        p {{ margin: 8px 0; }}
        b {{ color: #000; }}
        ul {{ margin: 5px 0; padding-left: 20px; }}
        li {{ margin-bottom: 4px; }}
        .image-container {{ text-align: center; margin: 20px 0; page-break-inside: avoid; }}
        img {{ border: 1px solid #ccc; border-radius: 4px; padding: 5px; background: #fff; max-width: 80%; height: auto; }}
        table {{ border-collapse: collapse; width: 100%; margin: 15px 0; table-layout: auto; page-break-inside: avoid; }}
        th, td {{ border: 1px solid #aaa; padding: 8px 12px; text-align: left; font-size: 13px; }}
        th {{ background-color: #e8eef5; font-weight: bold; color: #333; }}
        tr:nth-child(even) td {{ background-color: #f9f9f9; }}
        .visual-summary {{ background: #fdf6e3; border-left: 4px solid #b58900; padding: 12px 15px; font-size: 0.95em; color: #586e75; margin: 10px 0 20px 0; page-break-inside: avoid; }}
    </style>
</head>
<body>
    {body_html}
</body>
</html>"""

    def _apply_inline_styles(self, text: str) -> str:
        # PR #9: Use pre-compiled bold regex
        text = self.RE_BOLD.sub(r'<b>\1</b>', text)
        if "[画像概要]" in text:
            text = f'<div class="visual-summary">{text}</div>'
        return text

    def _render_table(self, rows: list) -> str:
        if not rows:
            return ""
        html_out = ["<table>"]
        for i, row in enumerate(rows):
            cells = [c.strip() for c in row.split("|")[1:-1]]
            tag = "th" if i == 0 else "td"
            html_out.append("<tr>")
            for cell in cells:
                # 🔒 Security Fix: HTML escape before applying inline styles
                escaped_cell = html.escape(cell)
                html_out.append(f"<{tag}>{self._apply_inline_styles(escaped_cell)}</{tag}>")
            html_out.append("</tr>")
        html_out.append("</table>")
        return "\n".join(html_out)

    def _resolve_images_to_tmpdir(self, md_content: str, tmp_dir: Path) -> str:
        def resolve_and_copy(match):
            rel_path = match.group(1)
            filename = Path(rel_path).name
            search_dirs = [self.output_dir / "media", self.output_dir]
            for search_dir in search_dirs:
                src = search_dir / filename
                if src.exists():
                    dst = tmp_dir / filename
                    shutil.copy2(str(src), str(dst))
                    return f'![画像](file://{dst.absolute()})'
            return match.group(0)
        
        return self.RE_IMAGE.sub(resolve_and_copy, md_content)

    def generate_pdf(self, md_content: str, output_name: str) -> Path:
        """Markdown内容からPDFを生成する (PR #22: Use secure tempfile)"""
        with tempfile.TemporaryDirectory(prefix="pdf_gen_") as tmp_dir_name:
            tmp_dir = Path(tmp_dir_name)
            md_content_resolved = self._resolve_images_to_tmpdir(md_content, tmp_dir)
            html_content = self._simple_md_to_html(md_content_resolved)
            
            temp_html = tmp_dir / f"{output_name}.html"
            with open(temp_html, "w", encoding="utf-8") as f:
                f.write(html_content)
            
            pdf_path = self.output_dir / f"{output_name}.pdf"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                # 🔒 Security Fix: Use absolute paths to prevent argument injection
                cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir.resolve()), str(temp_html.resolve())]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                if res.returncode != 0:
                    return None
                
                pdfs = list(tmp_dir.rglob("*.pdf"))
                if pdfs:
                    shutil.move(str(pdfs[0]), str(pdf_path))
                    return pdf_path
            except Exception:
                return None
        return None
