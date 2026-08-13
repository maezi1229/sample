"""パス設定。

作業中の一時ファイルはBox同期フォルダの外（ローカルディスク）に置くこと。
実運用時はこのファイルの値を各PCの実際のパスに合わせて変更してください。
"""
from pathlib import Path

# ローカル作業用フォルダ（Box同期対象外にすること）
LOCAL_WORK_DIR = Path(__file__).resolve().parent.parent / "work"

INPUT_DIR = LOCAL_WORK_DIR / "input_samples"
OCR_TEXT_DIR = LOCAL_WORK_DIR / "ocr_text"
DRAFT_DIR = LOCAL_WORK_DIR / "extracted_draft"
REVIEW_DIR = LOCAL_WORK_DIR / "review_sheets"

for _d in (LOCAL_WORK_DIR, INPUT_DIR, OCR_TEXT_DIR, DRAFT_DIR, REVIEW_DIR):
    _d.mkdir(parents=True, exist_ok=True)

TESSERACT_LANG = "jpn+eng"
OCR_DPI = 300
