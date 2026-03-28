import re
import subprocess
import shutil
import logging
import html
from pathlib import Path
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)

class DocumentGenerator:
    """MarkdownからHTML/PDFを生成するクラス"""
    
    RE_TABLE_SEP = re.compile(r'^:?-{2,}:?$')
    RE_HEADER = re.compile(r'^#+')
    RE_BOLD = re.compile(r'\*\*(.*?)\*\*')
    RE_IMAGE = re.compile(r'!\[.*?\]\((.*?)\)')

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir)

    def _apply_inline_styles(self, text: str) -> str:
        """太字等のインラインスタイルを適用する"""
        return self.RE_BOLD.sub(r'<b>\1</b>', text)

    def _render_table(self, lines: List[str]) -> str:
        """MarkdownのテーブルをHTMLに変換する"""
        if len(lines) < 2: return ""
        html_rows = ["<table border='1'>"]
        for i, line in enumerate(lines):
            if i == 1 and self.RE_TABLE_SEP.match(line.strip().split('|')[1] if '|' in line else ""):
                continue
            cells = [c.strip() for c in line.split('|') if c.strip()]
            tag = "th" if i == 0 else "td"
            row_content = "".join([f"<{tag}>{self._apply_inline_styles(html.escape(c))}</{tag}>" for c in cells])
            html_rows.append(f"  <tr>{row_content}</tr>")
        html_rows.append("</table>")
        return "\n".join(html_rows)

    def _render_header(self, stripped_line: str) -> str:
        """ヘッダー要素をレンダリングする"""
        level = len(self.RE_HEADER.match(stripped_line).group())
        content = stripped_line.lstrip('#').strip()
        escaped_content = html.escape(content)
        styled_content = self._apply_inline_styles(escaped_content)
        return f"<h{level}>{styled_content}</h{level}>"

    def _render_list_item(self, stripped_line: str) -> str:
        """リスト項目をレンダリングする"""
        content = stripped_line.lstrip('-*').strip()
        escaped_content = html.escape(content)
        styled_content = self._apply_inline_styles(escaped_content)
        return f"<li>{styled_content}</li>"

    def _render_image_element(self, stripped_line: str) -> str:
        """画像要素をレンダリングする"""
        img_match = self.RE_IMAGE.search(stripped_line)
        img_path = img_match.group(1)
        # 🔒 Security Fix: HTML escape image source attribute
        escaped_img_path = html.escape(img_path, quote=True)
        return f'<div class="image-container"><img src="{escaped_img_path}" alt="画像"></div>'

    def _render_paragraph(self, stripped_line: str) -> str:
        """段落要素をレンダリングする"""
        escaped_line = html.escape(stripped_line)
        styled_line = self._apply_inline_styles(escaped_line)
        return f"<p>{styled_line}</p>"

    def _simple_md_to_html(self, md_content: str) -> str:
        """簡易的なMarkdownをHTMLに変換する"""
        lines = md_content.split('\n')
        html_output = [
            "<html><head><meta charset='utf-8'><style>",
            "body { font-family: sans-serif; margin: 2em; }",
            "table { border-collapse: collapse; width: 100%; margin-bottom: 1em; }",
            "th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }",
            "th { background-color: #f2f2f2; }",
            ".image-container { text-align: center; margin: 1em 0; }",
            "img { max-width: 100%; height: auto; }",
            "</style></head><body>"
        ]
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            if not stripped:
                i += 1
                continue
            
            if stripped.startswith('|'):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i])
                    i += 1
                html_output.append(self._render_table(table_lines))
            elif self.RE_HEADER.match(stripped):
                html_output.append(self._render_header(stripped))
                i += 1
            elif stripped.startswith(('-', '*')):
                html_output.append("<ul>")
                while i < len(lines) and lines[i].strip().startswith(('-', '*')):
                    html_output.append(self._render_list_item(lines[i].strip()))
                    i += 1
                html_output.append("</ul>")
            elif stripped.startswith('!['):
                html_output.append(self._render_image_element(stripped))
                i += 1
            else:
                html_output.append(self._render_paragraph(stripped))
                i += 1
        
        html_output.append("</body></html>")
        return "\n".join(html_output)

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

    def generate_pdf(self, md_content: str, output_name: str) -> Optional[Path]:
        """MarkdownからPDFを生成する（LibreOffice sofficeを使用）"""
        import tempfile
        with tempfile.TemporaryDirectory() as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            html_content = self._simple_md_to_html(self._resolve_images_to_tmpdir(md_content, tmp_dir))
            
            temp_html = tmp_dir / f"{output_name}.html"
            with open(temp_html, "w", encoding="utf-8") as f:
                f.write(html_content)
            
            pdf_path = self.output_dir / f"{output_name}.pdf"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(temp_html)]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                if res.returncode != 0:
                    logger.error(f"soffice conversion failed (returncode {res.returncode}): {res.stderr}")
                    return None
                
                pdfs = list(tmp_dir.rglob("*.pdf"))
                if pdfs:
                    shutil.move(str(pdfs[0]), str(pdf_path))
                    return pdf_path
                else:
                    logger.error(f"soffice succeeded but no PDF was found in {tmp_dir}")
            except (subprocess.SubprocessError, OSError) as e:
                logger.error(f"Failed to generate PDF for {output_name}: {e}")
                return None
            except Exception:
                logger.exception("Unexpected error during PDF generation")
                return None
        return None
