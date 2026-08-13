"""PDF読み取りユーティリティ：テキスト抽出 / OCR。

仕入先PDFには「Excel等から直接出力されたテキスト付きPDF」と
「スキャン・FAXの画像PDF」が混在しうるため、まずPDF内蔵のテキスト層を試し、
文字数が少ない（＝画像PDF）場合のみOCRにフォールバックする。
"""
import sys
from pathlib import Path

import fitz  # pymupdf
import pytesseract
from PIL import Image
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation import config

if config.TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = config.TESSERACT_CMD

MIN_TEXT_LEN_FOR_NATIVE = 20


def extract_native_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def render_page_image(pdf_path: Path, page_index: int, dpi: int) -> Image.Image:
    doc = fitz.open(str(pdf_path))
    pix = doc[page_index].get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    doc.close()
    return img


def preprocess_for_ocr(img: Image.Image, threshold: int = 180) -> Image.Image:
    gray = img.convert("L")
    return gray.point(lambda x: 0 if x < threshold else 255, mode="1")


def ocr_page(pdf_path: Path, page_index: int, dpi: int, lang: str) -> str:
    img = render_page_image(pdf_path, page_index, dpi)
    bw = preprocess_for_ocr(img)
    return pytesseract.image_to_string(bw, lang=lang, config="--psm 6")


def read_pdf_text(pdf_path: Path, dpi: int = 300, lang: str = "jpn+eng") -> tuple[str, str]:
    """PDFからテキストを取得する。戻り値は (text, method)。methodは "native" か "ocr"。"""
    native = extract_native_text(pdf_path).strip()
    if len(native) >= MIN_TEXT_LEN_FOR_NATIVE:
        return native, "native"

    page_count = len(PdfReader(str(pdf_path)).pages)
    texts = [ocr_page(pdf_path, i, dpi, lang) for i in range(page_count)]
    return "\n".join(texts), "ocr"
