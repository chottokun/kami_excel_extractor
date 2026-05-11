import subprocess
import tempfile
import logging
import shutil
from pathlib import Path
from typing import Optional, List, Union

logger = logging.getLogger(__name__)

class ExcelConverter:
    """Excelを画像に変換するクラス (PDF経由)"""
    
    def __init__(self, output_dir: Path, dpi: int = 150):
        self.output_dir = Path(output_dir).resolve()
        self.dpi = dpi

    def convert(self, input_file: Path, sheet_name: Optional[str] = None) -> Union[Path, List[Path]]:
        import uuid
        import openpyxl
        input_file = input_file.resolve()
        
        # 🔒 Race Condition Fix: UUIDを使用して中間ファイル名の衝突を回避
        run_id = str(uuid.uuid4())[:8]
        output_prefix = self.output_dir / f"{input_file.stem}_{run_id}"
        original_pdf = self.output_dir / f"{input_file.stem}_{run_id}.pdf"

        # 入力ファイルの存在確認
        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        try:
            with tempfile.TemporaryDirectory(prefix="lo_profile_") as tmp_dir_str:
                tmp_dir = Path(tmp_dir_str).resolve()
                
                # ターゲットファイルの準備 (シート隔離が必要な場合)
                target_excel = input_file
                if sheet_name:
                    target_excel = tmp_dir / f"isolated_{run_id}.xlsx"
                    logger.info(f"Isolating sheet '{sheet_name}' for conversion...")
                    wb = openpyxl.load_workbook(input_file, data_only=True)
                    for name in wb.sheetnames:
                        if name != sheet_name:
                            del wb[name]
                    wb.save(target_excel)

                # Step 1: Excel -> PDF
                logger.info(f"Converting {target_excel.name} to PDF (ID: {run_id})...")
                
                raw_soffice_path = shutil.which("soffice")
                if not raw_soffice_path:
                    logger.error("LibreOffice (soffice) not found in PATH")
                    raise RuntimeError("LibreOffice (soffice) not found in PATH")
                soffice_path = str(Path(raw_soffice_path).resolve())

                # 🔒 Security Fix: Use absolute paths to prevent argument injection
                res_pdf = subprocess.run([
                    str(Path(soffice_path).resolve()),
                    f"-env:UserInstallation=file://{tmp_dir.resolve()}",
                    "--headless", "--convert-to", "pdf",
                    "--outdir", str(tmp_dir.resolve()),
                    str(target_excel.resolve())
                ], capture_output=True, text=True, timeout=600)

                if res_pdf.returncode != 0:
                    logger.error(f"LibreOffice failed: {res_pdf.stderr}")
                    raise RuntimeError(f"LibreOffice conversion failed: {res_pdf.stderr}")

                # LibreOfficeは元のファイル名でPDFを書き出すため、生成後にリネームする
                default_pdf = tmp_dir / f"{target_excel.stem}.pdf"
                if default_pdf.exists():
                    shutil.move(str(default_pdf), str(original_pdf))

                if not original_pdf.exists():
                    raise FileNotFoundError(f"PDF not found after conversion: {original_pdf}")

                # Step 2: PDF -> PNG (multi-page support)
                logger.info(f"Converting PDF to PNG (multi-page: {bool(sheet_name)})...")
                
                if sheet_name:
                    # シートごとのマルチページ抽出
                    return self._convert_pdf_to_multi_png(original_pdf, output_prefix)
                else:
                    # 全体概要用の単一ファイル抽出 (後方互換性)
                    output_png = output_prefix.with_suffix(".png")
                    self._convert_pdf_to_png(original_pdf, output_png)
                    return output_png
        finally:
            if original_pdf.exists():
                original_pdf.unlink()

    def _convert_pdf_to_multi_png(self, pdf_path: Path, output_prefix: Path) -> List[Path]:
        """PDFの全ページをPNGに変換する"""
        raw_path = shutil.which("pdftocairo")
        if not raw_path:
            # pdftocairo がない場合は fitz を使用 (フォールバック)
            return self._try_fitz_multi(pdf_path, output_prefix)

        try:
            # pdftocairo を使用して全ページを連番で出力
            subprocess.run([
                str(Path(raw_path).resolve()), "-png",
                str(pdf_path.resolve()), str(output_prefix.resolve())
            ], check=True, capture_output=True, timeout=300)
            
            # 生成されたファイルを収集 (prefix-1.png, prefix-2.png, ...)
            pngs = sorted(list(self.output_dir.glob(f"{output_prefix.name}-*.png")), 
                         key=lambda p: int(p.stem.split("-")[-1]))
            if pngs:
                logger.info(f"Generated {len(pngs)} PNGs using pdftocairo")
                return pngs
        except Exception as e:
            logger.warning(f"pdftocairo multi-page failed: {e}")

        return self._try_fitz_multi(pdf_path, output_prefix)

    def _try_fitz_multi(self, pdf_path: Path, output_prefix: Path) -> List[Path]:
        """PyMuPDFを使用して全ページをPNGに変換する"""
        try:
            import fitz
            doc = fitz.open(str(pdf_path.resolve()))
            pngs = []
            for i in range(len(doc)):
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                out_path = self.output_dir / f"{output_prefix.name}-{i+1}.png"
                pix.save(str(out_path.resolve()))
                pngs.append(out_path)
            doc.close()
            if pngs:
                logger.info(f"Generated {len(pngs)} PNGs using fitz")
                return pngs
        except Exception as e:
            logger.error(f"Multi-page conversion failed: {e}")
        
        raise RuntimeError("Multi-page PDF to PNG conversion failed")

    def _convert_pdf_to_png(self, pdf_path: Path, output_png: Path) -> None:
        """PDFをPNGに変換する (複数の方法を試行するフォールバックチェーン)"""

        # 1. pdftocairo (Primary)
        if self._try_pdftocairo(pdf_path, output_png):
            return

        # 2. PyMuPDF (fitz)
        if self._try_fitz(pdf_path, output_png):
            return

        # 3. ImageMagick (magick or convert)
        if self._try_imagemagick(pdf_path, output_png):
            return

        raise RuntimeError("All PDF to PNG conversion methods failed")

    def _try_pdftocairo(self, pdf_path: Path, output_png: Path) -> bool:
        """pdftocairoを使用してPDFをPNGに変換する"""
        raw_path = shutil.which("pdftocairo")
        if not raw_path:
            logger.warning("pdftocairo not found in PATH")
            return False

        try:
            # 🔒 Security Fix: Use absolute paths to prevent argument injection and CWE-426
            res = subprocess.run([
                str(Path(raw_path).resolve()), "-png", "-singlefile",
                str(pdf_path.resolve()), str(output_png.with_suffix("").resolve())
            ], capture_output=True, text=True, timeout=300)

            if res.returncode == 0 and output_png.exists():
                logger.info("Converted PDF to PNG using pdftocairo")
                return True
            logger.warning(f"pdftocairo failed: {res.stderr}")
        except (subprocess.SubprocessError, OSError) as e:
            logger.warning(f"pdftocairo failed: {e}")
        return False

    def _try_fitz(self, pdf_path: Path, output_png: Path) -> bool:
        """PyMuPDF (fitz) を使用してPDFをPNGに変換する"""
        try:
            import fitz
            doc = fitz.open(str(pdf_path.resolve()))
            page = doc.load_page(0)
            # 2.0x zoom for better quality
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            pix.save(str(output_png.resolve()))
            doc.close()
            if output_png.exists():
                logger.info("Converted PDF to PNG using PyMuPDF (fitz)")
                return True
        except ImportError:
            logger.warning("PyMuPDF (fitz) not installed")
        except Exception as e:
            logger.warning(f"PyMuPDF conversion failed: {e}")
        return False

    def _try_imagemagick(self, pdf_path: Path, output_png: Path) -> bool:
        """ImageMagick (magick or convert) を使用してPDFをPNGに変換する"""
        for cmd_name in ["magick", "convert"]:
            raw_cmd_path = shutil.which(cmd_name)
            if not raw_cmd_path:
                continue
            try:
                # 🔒 Security Fix: Use absolute paths to prevent argument injection
                # magick [input] [output] or convert [input] [output]
                # For PDF to PNG with ImageMagick, [0] specifies the first page
                res = subprocess.run([
                    str(Path(raw_cmd_path).resolve()), "-density", str(self.dpi),
                    f"{pdf_path.resolve()}[0]", str(output_png.resolve())
                ], capture_output=True, text=True, timeout=300)

                if res.returncode == 0 and output_png.exists():
                    logger.info(f"Converted PDF to PNG using ImageMagick ({cmd_name})")
                    return True
            except (subprocess.SubprocessError, OSError) as e:
                logger.warning(f"ImageMagick ({cmd_name}) failed: {e}")
                continue
        return False
