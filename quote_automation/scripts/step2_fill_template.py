"""Step2: レビューシートで確定した仕入先見積データを、自社見積書テンプレートに転記する。

Excelアプリは起動せず、openpyxlでテンプレートファイルを直接コピー・編集する。
入力はStep1で人が確認・修正済みのレビューシート(xlsx)のみを受け付ける
（確定値が空欄なら ReviewNotConfirmedError で止まる）。

使い方:
    python step2_fill_template.py <レビューシートxlsx> <テンプレートxlsx> <出力先xlsx>
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation.extraction.confirmed_reader import load_confirmed_supplier_quote
from quote_automation.template.fill_quote import fill_quote_template


def run(review_path: Path, template_path: Path, output_path: Path) -> Path:
    confirmed = load_confirmed_supplier_quote(review_path)

    print(f"確定済み見積書番号: {confirmed.quote_no}")
    print(f"確定済み日付: {confirmed.date}")
    print(f"確定済み合計金額: {confirmed.total_amount:,.0f}")
    print(f"明細: {len(confirmed.items)}件")
    for item in confirmed.items:
        print(f"  - {item.name} 数量={item.qty} 仕入単価={item.unit_price:,.0f}")

    items = [{"qty": i.qty, "unit_price": i.unit_price} for i in confirmed.items]
    result_path = fill_quote_template(template_path, output_path, items)
    print(f"転記済み見積書を出力しました: {result_path}")
    return result_path


def main():
    parser = argparse.ArgumentParser(description="レビューシートの確定値を自社見積書テンプレートへ転記する")
    parser.add_argument("review_path", type=Path)
    parser.add_argument("template_path", type=Path)
    parser.add_argument("output_path", type=Path)
    args = parser.parse_args()
    run(args.review_path, args.template_path, args.output_path)


if __name__ == "__main__":
    main()
