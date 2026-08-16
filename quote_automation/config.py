"""パス設定。

処理はすべてClaude Code on the webのクラウドセッション内で完結する
（会社PCには何もインストールしない）。work以下はそのセッションの中だけに
存在する作業用フォルダで、会社PCのBox同期フォルダとは無関係。
"""
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
LOCAL_WORK_DIR = PACKAGE_DIR.parent / "work"

INPUT_DIR = LOCAL_WORK_DIR / "input_samples"
OCR_TEXT_DIR = LOCAL_WORK_DIR / "ocr_text"
DRAFT_DIR = LOCAL_WORK_DIR / "extracted_draft"
CONFIRMED_DIR = LOCAL_WORK_DIR / "confirmed_quotes"
OUTPUT_DIR = LOCAL_WORK_DIR / "output"
OUTPUT_PDF_DIR = LOCAL_WORK_DIR / "output_pdf"

for _d in (LOCAL_WORK_DIR, INPUT_DIR, OCR_TEXT_DIR, DRAFT_DIR, CONFIRMED_DIR, OUTPUT_DIR, OUTPUT_PDF_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# チャット確認済みのデータを転記する、自社共通の空見積書テンプレート（リポジトリに常備）。
# 客先名・担当者名・件名・上乗せ率・備考は毎回 fill_quote_template() が書き込むため、
# 案件ごとにテンプレートを用意し直す必要はない。
BASE_TEMPLATE_PATH = PACKAGE_DIR / "assets" / "base_quote_template.xlsx"

# 見積り集計表（仕入金額・売り金額・利益率・受注状況など）。
# 客先名・金額など機密情報を含むため、Gitにはコミットせずwork/配下のみで管理する。
# セッションをまたいで蓄積したい場合は、都度ダウンロード→次回セッションで
# 同じ場所に再アップロードしてから使う。
LEDGER_PATH = LOCAL_WORK_DIR / "summary" / "quotation_ledger.xlsx"

TESSERACT_LANG = "jpn+eng"
OCR_DPI = 300
