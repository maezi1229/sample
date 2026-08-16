"""Step1: 仕入先PDF見積を読み取り、候補値をチャットで確認できる形で出力する。

Excelレビューシートは作らない（旧仕様は廃止）。ここで得られる値はあくまで
OCR/機械読み取りの「たたき台」であり、Claudeがこの結果をチャットで
ユーザーに提示し、不明点や誤読を1件ずつ確認・訂正してから
（客先名・担当者名・上乗せ率・備考の確認と合わせて）Step2に渡す。

使い方:
    python step1_extract_supplier_quote.py <仕入先PDFのパス>
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation import config
from quote_automation.extraction.field_parser import parse_draft
from quote_automation.extraction.pdf_reader import read_pdf_text


def run(pdf_path: Path) -> Path:
    text, method = read_pdf_text(pdf_path, dpi=config.OCR_DPI, lang=config.TESSERACT_LANG)
    draft = parse_draft(text)
    stem = pdf_path.stem

    (config.OCR_TEXT_DIR / f"{stem}.txt").write_text(text, encoding="utf-8")

    draft_json = {
        "source_pdf": str(pdf_path),
        "method": method,
        "quote_no": draft.quote_no,
        "date": draft.date,
        "total_amount": draft.total_amount,
        "line_items": [item.__dict__ for item in draft.line_items],
    }
    draft_path = config.DRAFT_DIR / f"{stem}_draft.json"
    draft_path.write_text(json.dumps(draft_json, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"読み取り方式: {method}")
    print(f"見積書番号: {draft.quote_no!r}")
    print(f"日付: {draft.date!r}")
    print(f"見積金額: {draft.total_amount!r}")
    print(f"明細候補: {len(draft.line_items)}件")
    for item in draft.line_items:
        print(f"  - {item.name_guess!r} 数量={item.qty_guess} 単価={item.unit_price_guess} 金額={item.amount_guess}")
    print(f"OCRテキスト出力先: {config.OCR_TEXT_DIR / f'{stem}.txt'}")
    print(f"抽出下書きJSON出力先: {draft_path}")
    print()
    print("※この結果はあくまで候補です。チャット上で元PDFと見比べ、"
          "数字・客先名・担当者名・上乗せ率・備考を確認してから見積書を作成してください。")

    return draft_path


def main():
    parser = argparse.ArgumentParser(description="仕入先PDF見積を読み取り、候補値をJSON/テキストで出力する")
    parser.add_argument("pdf_path", type=Path, help="仕入先見積PDFのパス")
    args = parser.parse_args()
    run(args.pdf_path)


if __name__ == "__main__":
    main()
