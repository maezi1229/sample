"""パス設定。

【一度だけ、各PCに合わせて確認・変更してください】
- LOCAL_WORK_DIR: 作業中の一時ファイルの保存先。Box同期フォルダの外（ローカルディスク）を指定すること。
  例: Path(r"C:\\Users\\your_name\\quote_work")
- TESSERACT_CMD: WindowsでTesseract OCRをインストールした場所。
  インストーラでPATHに追加していない場合はここにフルパスを指定する。
  例: r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
  PATHに追加済み、またはMac/Linuxの場合は None のままでよい。
"""
from pathlib import Path

# ローカル作業用フォルダ（Box同期対象外にすること）
LOCAL_WORK_DIR = Path(__file__).resolve().parent.parent / "work"

TESSERACT_CMD = None  # 例: r"C:\Program Files\Tesseract-OCR\tesseract.exe"

INPUT_DIR = LOCAL_WORK_DIR / "input_samples"
OCR_TEXT_DIR = LOCAL_WORK_DIR / "ocr_text"
DRAFT_DIR = LOCAL_WORK_DIR / "extracted_draft"
REVIEW_DIR = LOCAL_WORK_DIR / "review_sheets"

for _d in (LOCAL_WORK_DIR, INPUT_DIR, OCR_TEXT_DIR, DRAFT_DIR, REVIEW_DIR):
    _d.mkdir(parents=True, exist_ok=True)

TESSERACT_LANG = "jpn+eng"
OCR_DPI = 300
