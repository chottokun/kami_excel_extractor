import re
import subprocess
import shutil
import logging
import html
from pathlib import Path
from typing import List, Optional
from .utils import secure_filename

logger = logging.getLogger(__name__)

class DocumentGenerator:
    """MarkdownからHTML/PDFを生成するクラス"""
    
    RE_TABLE_SEP = re.compile(r'^:?-{2,}:?$')
    RE_HEADER = re.compile(r'^#+')
    RE_BOLD = re.compile(r'\*\*(.*?)\*\*')
    RE_IMAGE = re.compile(r'!\[(.*?)\]\((.*?)\)')

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir).resolve()

    def _render_inline(self, text: str) -> str:
        """テキストをHTMLエスケープし、インラインスタイルを適用する"""
        # 🔒 Security Fix: HTML escape before applying inline styles
        return self._apply_inline_styles(html.escape(text))

    def _apply_inline_styles(self, text: str) -> str:
        """太字等のインラインスタイルを適用する"""
        text = self.RE_BOLD.sub(r'<b>\1</b>', text)
        if "[画像概要]" in text:
            text = f'<div class="visual-summary">{text}</div>'
        return text

    def _render_table(self, lines: List[str]) -> str:
        """MarkdownのテーブルをHTMLに変換する"""
        if not lines:
            return ""
        html_out = ["<table>"]
        for i, line in enumerate(lines):
            # セパレータ行をスキップ
            if i == 1 and self.RE_TABLE_SEP.match(line.strip().split('|')[1] if '|' in line else ""):
                continue
            
            # セルを分割。先頭と末尾の空要素を考慮
            cells = [c.strip() for c in line.split('|') if c.strip()]
            tag = "th" if i == 0 else "td"
            html_out.append("  <tr>")
            for cell in cells:
                html_out.append(f"    <{tag}>{self._render_inline(cell)}</{tag}>")
            html_out.append("  </tr>")
        html_out.append("</table>")
        return "\n".join(html_out)

    def _render_header(self, stripped_line: str) -> str:
        """ヘッダー要素をレンダリングする"""
        header_match = self.RE_HEADER.match(stripped_line)
        level = len(header_match.group()) if header_match else 1
        content = stripped_line.lstrip("#").strip()
        return f"<h{level}>{self._render_inline(content)}</h{level}>"

    def _render_list_item(self, stripped_line: str) -> str:
        """リストアイテムをレンダリングする"""
        # 行頭の - または * とその後のスペースを除去 (正規表現で確実に分離)
        match = re.match(r'^[-*]\s+(.*)$', stripped_line)
        if match:
            content = match.group(1)
        else:
            # スペースがない場合は記号1文字だけ削る
            content = stripped_line[1:].strip()
        return f"<li>{self._render_inline(content)}</li>"

    def _render_image_element(self, stripped_line: str) -> str:
        """画像要素をレンダリングする"""
        img_match = self.RE_IMAGE.search(stripped_line)
        alt_text = img_match.group(1) or "画像"
        img_path = img_match.group(2)
        # 🔒 Security Fix: HTML escape image source attribute
        escaped_img_path = html.escape(img_path, quote=True)
        escaped_alt = html.escape(alt_text)
        return f'<div class="image-container"><img src="{escaped_img_path}" alt="{escaped_alt}"></div>'

    def _render_paragraph(self, stripped_line: str) -> str:
        """段落要素をレンダリングする"""
        return f"<p>{self._render_inline(stripped_line)}</p>"

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

    def _simple_md_to_html(self, md_content: str) -> str:
        """簡易的なMarkdownをHTMLに変換する"""
        lines = md_content.split('\n')
        body_parts = []
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            if not stripped:
                i += 1
                continue
            
            elif stripped.startswith('|'):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i])
                    i += 1
                body_parts.append(self._render_table(table_lines))
            elif self.RE_HEADER.match(stripped):
                body_parts.append(self._render_header(stripped))
                i += 1
            elif re.match(r'^[-*](\s+|$)', stripped):
                body_parts.append("<ul>")
                while i < len(lines) and re.match(r'^[-*](\s+|$)', lines[i].strip()):
                    body_parts.append(self._render_list_item(lines[i].strip()))
                    i += 1
                body_parts.append("</ul>")
            elif stripped.startswith('!['):

                body_parts.append(self._render_image_element(stripped))
                i += 1
            else:
                body_parts.append(self._render_paragraph(stripped))
                i += 1
        
        return self._get_html_template("\n".join(body_parts))

    def _resolve_images_to_tmpdir(self, md_content: str, tmp_dir: Path) -> str:
        def resolve_and_copy(match):
            alt_text = match.group(1)
            rel_path = match.group(2)
            filename = Path(rel_path).name
            search_dirs = [self.output_dir / "media", self.output_dir]
            for search_dir in search_dirs:
                src = search_dir / filename
                if src.exists():
                    dst = tmp_dir / filename
                    shutil.copy2(str(src), str(dst))
                    return f'![{alt_text}](file://{dst.absolute()})'
            return match.group(0)
        
        return self.RE_IMAGE.sub(resolve_and_copy, md_content)

    def generate_pdf(self, md_content: str, output_name: str) -> Optional[Path]:
        """MarkdownからPDFを生成する（LibreOffice sofficeを使用）"""
        import tempfile
        # 安全なファイル名を作成
        safe_output_name = secure_filename(output_name)
        
        with tempfile.TemporaryDirectory(prefix="pdf_gen_") as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str).resolve()
            html_content = self._simple_md_to_html(self._resolve_images_to_tmpdir(md_content, tmp_dir))
            
            # 一時ディレクトリ内ではフラットに管理
            temp_html = (tmp_dir / f"{safe_output_name}.html").resolve()
            with open(temp_html, "w", encoding="utf-8") as f:
                f.write(html_content)
            
            # 最終的な出力パス
            pdf_path = self.output_dir / f"{safe_output_name}.pdf"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            
            try:
                # 🔒 Security Fix: Use absolute path for executable to prevent untrusted search path (CWE-426)
                soffice_path = shutil.which("soffice")
                if not soffice_path:
                    logger.error("LibreOffice (soffice) not found in PATH")
                    return None

                # 🔒 Security Fix: Use absolute paths to prevent argument injection
                # --outdir は一時ディレクトリのルートを指定
                cmd = [soffice_path, "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(temp_html)]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                if res.returncode != 0:
                    logger.error(f"soffice conversion failed (returncode {res.returncode}): {res.stderr}")
                    return None
                
                # 生成されたPDFを特定（sofficeは入力ファイル名.pdfを出力する）
                expected_pdf = tmp_dir / f"{temp_html.stem}.pdf"
                if expected_pdf.exists():
                    shutil.move(str(expected_pdf), str(pdf_path))
                    return pdf_path
                else:
                    # フォールバック: rglobで探す
                    pdfs = list(tmp_dir.rglob("*.pdf"))
                    if pdfs:
                        shutil.move(str(pdfs[0]), str(pdf_path))
                        return pdf_path
                    logger.error(f"soffice succeeded but no PDF was found in {tmp_dir}")
            except (subprocess.SubprocessError, OSError) as e:
                logger.error(f"Failed to generate PDF for {output_name}: {e}")
                return None
            except Exception:
                logger.exception("Unexpected error during PDF generation")
                return None
        return None
