"""見積書xlsxをPDF化する。

LibreOffice Calc(headless)でxlsx→PDF変換を行う。Excelアプリは起動せず、
マクロも使わない。印刷範囲・用紙設定はテンプレート側(印刷範囲・拡大縮小印刷等)に
従うため、Excelで印刷したときと同じレイアウトになる。
"""
import subprocess
from pathlib import Path


class PdfExportError(Exception):
    pass


def export_to_pdf(xlsx_path: Path, output_dir: Path, profile_dir: Path) -> Path:
    xlsx_path = Path(xlsx_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(xlsx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise PdfExportError(f"PDF変換に失敗しました: {result.stdout}\n{result.stderr}")

    pdf_path = output_dir / (xlsx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise PdfExportError(f"PDFファイルが見つかりません: {pdf_path}")
    return pdf_path
