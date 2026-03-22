import subprocess
<<<<<<< HEAD
=======
import os
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
import tempfile
from pathlib import Path
import re
import logging
import shutil

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
                if in_list:
                    html_output.append("</ul>")
                    in_list = False
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
                if in_list:
                    html_output.append("</ul>")
                    in_list = False
                continue
                
            # Headers
            if stripped.startswith("#"):
                if in_list:
                    html_output.append("</ul>")
                    in_list = False
                header_match = self.RE_HEADER.match(stripped)
                level = len(header_match.group()) if header_match else 1
                content = stripped.lstrip("#").strip()
                html_output.append(f"<h{level}>{self._apply_inline_styles(content)}</h{level}>")
            # Lists
            elif stripped.startswith("- "):
                if not in_list:
                    html_output.append("<ul>")
                    in_list = True
                content = stripped[2:].strip()
                html_output.append(f"<li>{self._apply_inline_styles(content)}</li>")
            # Images
            elif self.RE_IMAGE.match(stripped):
                if in_list:
                    html_output.append("</ul>")
                    in_list = False
                img_match = self.RE_IMAGE.search(stripped)
                img_path = img_match.group(1)
                html_output.append(f'<div class="image-container"><img src="{img_path}" alt="画像"></div>')
            # Text
            else:
                if in_list:
                    html_output.append("</ul>")
                    in_list = False
                html_output.append(f"<p>{self._apply_inline_styles(stripped)}</p>")

        # 末尾処理
        if in_table:
            html_output.append(self._render_table(table_rows))
        if in_list:
            html_output.append("</ul>")

        template = f"""<!DOCTYPE html>
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
    {"\n".join(html_output)}
</body>
</html>"""
        return template

    def _apply_inline_styles(self, text: str) -> str:
        # PR #9: Use pre-compiled bold regex
        text = self.RE_BOLD.sub(r'<b>\1</b>', text)
        if "[画像概要]" in text:
            text = f'<div class="visual-summary">{text}</div>'
        return text

    def _render_table(self, rows: list) -> str:
        if not rows:
            return ""
        html = ["<table>"]
        for i, row in enumerate(rows):
            cells = [c.strip() for c in row.split("|")[1:-1]]
            tag = "th" if i == 0 else "td"
            html.append("<tr>")
            for cell in cells:
                html.append(f"<{tag}>{self._apply_inline_styles(cell)}</{tag}>")
            html.append("</tr>")
        html.append("</table>")
        return "\n".join(html)

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
<<<<<<< HEAD
        """Markdown内容からPDFを生成する"""
        # Permission問題を避けるため、作業は一時ディレクトリで行う
        with tempfile.TemporaryDirectory(prefix="pdf_gen_") as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)

            # 画像をtmp_dirにコピーし、パスを書き換え
            md_content_resolved = self._resolve_images_to_tmpdir(md_content, tmp_dir)

            # Markdown→HTML変換
            html_content = self._simple_md_to_html(md_content_resolved)

            temp_html = tmp_dir / f"{output_name}.html"
            temp_html.parent.mkdir(parents=True, exist_ok=True)

            pdf_path = self.output_dir / f"{output_name}.pdf"
            pdf_path.parent.mkdir(parents=True, exist_ok=True)

            # HTMLを保存
            with open(temp_html, "w", encoding="utf-8") as f:
                f.write(html_content)

            logger.info(f"Generated HTML at {temp_html}, size: {temp_html.stat().st_size} bytes")

            try:
                # sofficeを使ってPDF変換
                cmd = [
                    "soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", str(tmp_dir), str(temp_html)
                ]

                logger.info(f"PDF conversion cmd: {' '.join(cmd)}")
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

                if res.returncode != 0:
                    logger.error(f"soffice failed: stdout={res.stdout}, stderr={res.stderr}")
                    return None

                # 生成されたPDFを探す
                pdfs = list(tmp_dir.rglob("*.pdf"))
                if pdfs:
                    generated_pdf = pdfs[0]
                    shutil.move(str(generated_pdf), str(pdf_path))
                    logger.info(f"Successfully generated PDF: {pdf_path}")
                    return pdf_path
                else:
                    logger.error(f"No PDF found in {tmp_dir}. Contents: {list(tmp_dir.glob('*'))}")
                    return None
            except Exception as e:
                logger.error(f"Error during PDF generation: {e}")
                return None
=======
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
                cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp_dir), str(temp_html)]
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
>>>>>>> 7c30b98 (feat: integrate all improvements (Async, RAG optimization, Security, Image/JSON robustness) and cleanup redundant scripts)
