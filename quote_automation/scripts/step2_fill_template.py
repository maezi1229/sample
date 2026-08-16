"""Step2: チャットで確認済みの見積データ(JSON)を自社見積書テンプレートに転記する。

Excelアプリは起動せず、openpyxlでテンプレートファイルを直接コピー・編集する。
入力は、Claudeがチャットでの確認を経て作成した「確定済み見積JSON」のみを受け付ける
（必須項目が欠けていれば InvalidConfirmedQuoteError で止まる）。

出力ファイル名は自動的に「【見積書】<件名>.xlsx」になる。

見積書を1件作成するたびに、見積り集計表（work/summary/quotation_ledger.xlsx、
Gitにはコミットしない機密ファイル）にも自動で1行追記する
（--skip-ledger を指定すると追記しない。テスト実行時などに使う）。

使い方:
    python step2_fill_template.py <確定済み見積JSON> <テンプレートxlsx> <出力先フォルダ>
"""
import argparse
import datetime
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation import config
from quote_automation.extraction.confirmed_quote import load_confirmed_quote
from quote_automation.scripts.update_ledger import append_ledger_entry
from quote_automation.template.fill_quote import compute_customer_unit_price, fill_quote_template

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def build_output_filename(item_title: str) -> str:
    safe_title = INVALID_FILENAME_CHARS.sub("_", item_title).strip()
    return f"【見積書】{safe_title}.xlsx"


def run(confirmed_json_path: Path, template_path: Path, output_dir: Path, skip_ledger: bool = False) -> Path:
    confirmed = load_confirmed_quote(confirmed_json_path)

    print(f"客先名: {confirmed.customer_name}")
    print(f"担当者名: {confirmed.contact_name}")
    print(f"件名: {confirmed.item_title}")
    print(f"上乗せ率: {confirmed.markup_percent}%（1000円単位で切り上げ）")
    print(f"明細: {len(confirmed.items)}件")
    for item in confirmed.items:
        print(f"  - {item.name} 数量={item.qty} 仕入単価={item.unit_price:,.0f}")
    print(f"備考: {len(confirmed.remarks_lines)}行")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / build_output_filename(confirmed.item_title)

    items = [{"name": i.name, "qty": i.qty, "unit_price": i.unit_price} for i in confirmed.items]
    result_path = fill_quote_template(
        template_path,
        output_path,
        customer_name=confirmed.customer_name,
        contact_name=confirmed.contact_name,
        item_title=confirmed.item_title,
        markup_percent=confirmed.markup_percent,
        items=items,
        remarks_lines=confirmed.remarks_lines,
    )
    print(f"転記済み見積書を出力しました: {result_path}")

    if not skip_ledger:
        cost_amount = sum(i.qty * i.unit_price for i in confirmed.items)
        sell_amount = sum(
            i.qty * compute_customer_unit_price(i.unit_price, confirmed.markup_percent)
            for i in confirmed.items
        )
        profit_amount = sell_amount - cost_amount
        ledger_path = append_ledger_entry(
            config.LEDGER_PATH,
            date=datetime.date.today().strftime("%Y/%m/%d"),
            customer_name=confirmed.customer_name,
            contact_name=confirmed.contact_name,
            supplier_name=confirmed.supplier_name,
            supplier_contact=confirmed.supplier_contact,
            item_title=confirmed.item_title,
            cost_amount=cost_amount,
            sell_amount=sell_amount,
            profit_rate=confirmed.markup_percent,
            profit_amount=profit_amount,
        )
        print(f"見積り集計表に追記しました: {ledger_path}")

    return result_path


def main():
    parser = argparse.ArgumentParser(description="確定済み見積JSONを自社見積書テンプレートへ転記する")
    parser.add_argument("confirmed_json_path", type=Path)
    parser.add_argument("template_path", type=Path, nargs="?", default=config.BASE_TEMPLATE_PATH)
    parser.add_argument("output_dir", type=Path, nargs="?", default=config.OUTPUT_DIR)
    parser.add_argument("--skip-ledger", action="store_true", help="見積り集計表への追記を行わない（テスト実行時などに使用）")
    args = parser.parse_args()
    run(args.confirmed_json_path, args.template_path, args.output_dir, skip_ledger=args.skip_ledger)


if __name__ == "__main__":
    main()
