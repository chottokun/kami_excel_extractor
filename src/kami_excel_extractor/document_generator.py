import asyncio
import concurrent.futures
import re
import subprocess
import shutil
import logging
import html
from pathlib import Path
from typing import List, Optional, Callable, Awaitable, Union
from .utils import secure_filename

logger = logging.getLogger(__name__)

class DocumentGenerator:
    """MarkdownからHTML/PDFを生成するクラス"""
    
    RE_TABLE_SEP = re.compile(r'^:?-{2,}:?$')
    RE_HEADER = re.compile(r'^#+')
    RE_BOLD = re.compile(r'\*\*(.*?)\*\*')
    RE_LIST_ITEM_START = re.compile(r'^[-*](\s+|$)')
    RE_LIST_ITEM_CONTENT = re.compile(r'^[-*]\s+(.*)$')

    def __init__(self, output_dir: Path):
        self.output_dir = Path(output_dir).resolve()

    def _parse_balanced_image(self, line: str) -> Optional[tuple[str, str]]:
        """
        ![alt](path) 形式のMarkdownからaltとpathを抽出する。
        括弧のネストを正しく扱う。
        """
        stripped = line.strip()
        if not (stripped.startswith("![") and "]" in stripped and "(" in stripped):
            return None

        # altテキストの抽出
        alt_start = 2
        alt_end = stripped.find("]")
        if alt_end == -1: return None
        alt_text = stripped[alt_start:alt_end]

        # パスの抽出 (括弧のバランスを考慮)
        remaining = stripped[alt_end+1:]
        if not remaining.startswith("("):
            return None
        
        path_start = 1
        stack = 0
        for i, char in enumerate(remaining):
            if char == '(':
                stack += 1
            elif char == ')':
                stack -= 1
                if stack == 0:
                    path_content = remaining[path_start:i]
                    return alt_text, path_content
        return None

    def _render_inline(self, text: str) -> str:
        """テキストをHTMLエスケープし、インラインスタイル（画像、太字）を適用する"""
        
        # 1. 画像の抽出と置換 (エスケープ前に行う)
        result_parts = []
        last_end = 0
        i = 0
        while i < len(text):
            if text[i:i+2] == "![":
                parsed = self._parse_balanced_image(text[i:])
                if parsed:
                    # 画像の前のテキストをエスケープして追加
                    before_text = text[last_end:i]
                    result_parts.append(self._apply_inline_styles(html.escape(before_text, quote=True)))
                    
                    alt_text, path_content = parsed
                    escaped_img_path = html.escape(path_content, quote=True)
                    escaped_alt = html.escape(alt_text or "画像", quote=True)
                    img_tag = f'<div class="image-container"><img src="{escaped_img_path}" alt="{escaped_alt}"></div>'
                    result_parts.append(img_tag)
                    
                    consumed_len = len(f"![{alt_text}]({path_content})")
                    i += consumed_len
                    last_end = i
                    continue
            i += 1
        
        # 残りのテキストを追加
        remaining_text = text[last_end:]
        result_parts.append(self._apply_inline_styles(html.escape(remaining_text, quote=True)))
        
        return "".join(result_parts)

    def _apply_inline_styles(self, text: str) -> str:
        """太字等のインラインスタイルを適用する (textはエスケープ済み)"""
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
        match = self.RE_LIST_ITEM_CONTENT.match(stripped_line)
        if match:
            content = match.group(1)
        else:
            # スペースがない場合は記号1文字だけ削る
            content = stripped_line[1:].strip()
        return f"<li>{self._render_inline(content)}</li>"

    def _render_image_element(self, stripped_line: str) -> str:
        """
        画像要素をレンダリングする。
        括弧のネストを正しく扱うため、手動でバランスパースを行う。
        """
        # 形式: ![alt](path)
        if not (stripped_line.startswith("![") and "]" in stripped_line and "(" in stripped_line):
            return stripped_line

        try:
            # altテキストの抽出
            alt_start = 2
            alt_end = stripped_line.find("]")
            if alt_end == -1: return stripped_line
            alt_text = stripped_line[alt_start:alt_end]

            # パスの抽出 (括弧のバランスを考慮)
            # ] の直後が ( であることを確認
            remaining = stripped_line[alt_end+1:]
            if not remaining.startswith("("):
                return stripped_line
            
            path_start = 1 # remaining における開始位置
            stack = 0
            path_content = None
            
            for i, char in enumerate(remaining):
                if char == '(':
                    stack += 1
                elif char == ')':
                    stack -= 1
                    if stack == 0:
                        path_content = remaining[path_start:i]
                        break
            
            if path_content is None:
                return stripped_line

            # 🔒 Security Fix: HTML escape image source attribute and alt text
            escaped_img_path = html.escape(path_content, quote=True)
            escaped_alt = html.escape(alt_text or "画像", quote=True)
            return f'<div class="image-container"><img src="{escaped_img_path}" alt="{escaped_alt}"></div>'
            
        except Exception:
            # 万が一のパース失敗時は元のテキストを返す
            return stripped_line

    def _render_paragraph(self, stripped_line: str) -> str:
        """段落要素をレンダリングする"""
        return f"<p>{self._render_inline(stripped_line)}</p>"

    def _process_table_block(self, lines: List[str], index: int) -> tuple[str, int]:
        """テーブルブロックを抽出してHTMLに変換する"""
        table_lines = []
        while index < len(lines) and lines[index].strip().startswith('|'):
            table_lines.append(lines[index])
            index += 1
        return self._render_table(table_lines), index

    def _process_list_block(self, lines: List[str], index: int) -> tuple[str, int]:
        """リストブロックを抽出してHTMLに変換する"""
        list_parts = ["<ul>"]
        while index < len(lines) and self.RE_LIST_ITEM_START.match(lines[index].strip()):
            list_parts.append(self._render_list_item(lines[index].strip()))
            index += 1
        list_parts.append("</ul>")
        return "\n".join(list_parts), index

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
            elif stripped.startswith('|'):
                html_table, i = self._process_table_block(lines, i)
                body_parts.append(html_table)
            elif self.RE_HEADER.match(stripped):
                body_parts.append(self._render_header(stripped))
                i += 1
            elif self.RE_LIST_ITEM_START.match(stripped):
                html_list, i = self._process_list_block(lines, i)
                body_parts.append(html_list)
            else:
                body_parts.append(self._render_paragraph(stripped))
                i += 1

        return self._get_html_template("\n".join(body_parts))

    def _resolve_single_image(self, line: str, search_dirs: List[Path], tmp_dir: Path) -> str:
        """単一の画像タグを解決し、ファイルをコピーする（同期版）"""
        parsed = self._parse_balanced_image(line)
        if not parsed:
            return line

        alt_text, rel_path = parsed
        filename = Path(rel_path).name
        for search_dir in search_dirs:
            src = (search_dir / filename).resolve()
            if src.exists():
                dst = (tmp_dir / filename).resolve()
                shutil.copy2(str(src), str(dst))
                return f'![{alt_text}](file://{dst})'
        return line

    async def _aresolve_single_image(self, line: str, search_dirs: List[Path], tmp_dir: Path) -> str:
        """単一の画像タグを解決し、ファイルをコピーする（非同期版）"""
        parsed = self._parse_balanced_image(line)
        if not parsed:
            return line

        alt_text, rel_path = parsed
        filename = Path(rel_path).name
        for search_dir in search_dirs:
            src = (search_dir / filename).resolve()
            # asyncio.to_thread を使用してブロッキングI/Oをスレッドで実行
            exists = await asyncio.to_thread(src.exists)
            if exists:
                dst = (tmp_dir / filename).resolve()
                await asyncio.to_thread(shutil.copy2, str(src), str(dst))
                return f'![{alt_text}](file://{dst})'
        return line

    def _resolve_images_to_tmpdir(self, md_content: str, tmp_dir: Path) -> str:
        """Markdown内の画像パスを絶対パス(file://)に変換し、ファイルを一時ディレクトリにコピーする (ThreadPoolExecutorで並列化)"""
        search_dirs = [self.output_dir / "media", self.output_dir]
        lines = md_content.splitlines()

        # 画像タグを含む行のみを並列処理の対象とする
        image_indices = [i for i, line in enumerate(lines) if "![" in line]
        if not image_indices:
            return md_content

        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = {executor.submit(self._resolve_single_image, lines[i], search_dirs, tmp_dir): i for i in image_indices}
            for future in concurrent.futures.as_completed(futures):
                i = futures[future]
                lines[i] = future.result()

        return "\n".join(lines)

    async def _aresolve_images_to_tmpdir(self, md_content: str, tmp_dir: Path) -> str:
        """Markdown内の画像パスを絶対パス(file://)に変換し、ファイルを一時ディレクトリにコピーする (asyncio.gatherで並列化)"""
        search_dirs = [self.output_dir / "media", self.output_dir]
        lines = md_content.splitlines()

        # 画像タグを含む行のみを並列処理の対象とする
        tasks = []
        task_indices = []
        for i, line in enumerate(lines):
            if "![" in line:
                tasks.append(self._aresolve_single_image(line, search_dirs, tmp_dir))
                task_indices.append(i)

        if not tasks:
            return md_content

        results = await asyncio.gather(*tasks)
        for idx, result in zip(task_indices, results):
            lines[idx] = result

        return "\n".join(lines)

    def _run_soffice_conversion(self, tmp_dir: Path, temp_html: Path) -> Optional[Path]:
        """LibreOfficeを使用してHTMLをPDFに変換し、結果のパスを返す"""
        try:
            # 🔒 Security Fix: Use absolute path for executable to prevent untrusted search path (CWE-426)
            raw_path = shutil.which("soffice")
            if not raw_path:
                logger.error("LibreOffice (soffice) not found in PATH")
                return None
            soffice_path = str(Path(raw_path).resolve())

            # 🔒 Security Fix: Use absolute paths to prevent argument injection
            # --outdir は一時ディレクトリのルートを指定
            res = subprocess.run([
                str(Path(soffice_path).resolve()), "--headless", "--convert-to", "pdf",
                "--outdir", str(tmp_dir.resolve()), str(temp_html.resolve())
            ], capture_output=True, text=True, timeout=60)
            if res.returncode != 0:
                logger.error(f"soffice conversion failed (returncode {res.returncode}): {res.stderr}")
                return None

            # 生成されたPDFを特定（sofficeは入力ファイル名.pdfを出力する）
            expected_pdf = tmp_dir / f"{temp_html.stem}.pdf"
            if expected_pdf.exists():
                return expected_pdf.resolve()

            # フォールバック: rglobで探す
            pdfs = list(tmp_dir.rglob("*.pdf"))
            if pdfs:
                return pdfs[0].resolve()

            logger.error(f"soffice succeeded but no PDF was found in {tmp_dir}")
            return None
        except (subprocess.SubprocessError, OSError) as e:
            logger.error(f"Subprocess error during soffice conversion: {e}")
            return None

    def _generate_pdf_internal(
        self,
        md_content: str,
        output_name: str,
        resolve_func: Union[Callable[[str, Path], str], Callable[[str, Path], Awaitable[str]]],
        is_async: bool = False
    ) -> Union[Optional[Path], Awaitable[Optional[Path]]]:
        """PDF生成の共通ロジック"""
        import tempfile
        safe_output_name = secure_filename(output_name)

        async def _async_logic():
            with tempfile.TemporaryDirectory(prefix="pdf_gen_") as tmp_dir_str:
                tmp_dir = Path(tmp_dir_str).resolve()
                resolved_md = await resolve_func(md_content, tmp_dir)
                html_content = self._simple_md_to_html(resolved_md)
                temp_html = (tmp_dir / f"{safe_output_name}.html").resolve()

                await asyncio.to_thread(lambda: temp_html.write_text(html_content, encoding="utf-8"))

                pdf_path = (self.output_dir / f"{safe_output_name}.pdf").resolve()
                await asyncio.to_thread(lambda: pdf_path.parent.mkdir(parents=True, exist_ok=True))

                try:
                    generated_pdf = await asyncio.to_thread(self._run_soffice_conversion, tmp_dir, temp_html)
                    if generated_pdf and generated_pdf.exists():
                        await asyncio.to_thread(shutil.move, str(generated_pdf.resolve()), str(pdf_path))
                        return pdf_path
                except Exception:
                    logger.exception("Unexpected error during async PDF generation")
                    return None
            return None

        def _sync_logic():
            with tempfile.TemporaryDirectory(prefix="pdf_gen_") as tmp_dir_str:
                tmp_dir = Path(tmp_dir_str).resolve()
                resolved_md = resolve_func(md_content, tmp_dir)
                html_content = self._simple_md_to_html(resolved_md)
                temp_html = (tmp_dir / f"{safe_output_name}.html").resolve()
                temp_html.write_text(html_content, encoding="utf-8")

                pdf_path = (self.output_dir / f"{safe_output_name}.pdf").resolve()
                pdf_path.parent.mkdir(parents=True, exist_ok=True)

                try:
                    generated_pdf = self._run_soffice_conversion(tmp_dir, temp_html)
                    if generated_pdf and generated_pdf.exists():
                        shutil.move(str(generated_pdf.resolve()), str(pdf_path))
                        return pdf_path
                except Exception:
                    logger.exception("Unexpected error during PDF generation")
                    return None
            return None

        return _async_logic() if is_async else _sync_logic()

    def generate_pdf(self, md_content: str, output_name: str) -> Optional[Path]:
        """MarkdownからPDFを生成する（LibreOffice sofficeを使用）"""
        return self._generate_pdf_internal(md_content, output_name, self._resolve_images_to_tmpdir, is_async=False)

    async def agenerate_pdf(self, md_content: str, output_name: str) -> Optional[Path]:
        """MarkdownからPDFを生成する（非同期版、LibreOffice sofficeを使用）"""
        return await self._generate_pdf_internal(md_content, output_name, self._aresolve_images_to_tmpdir, is_async=True)
