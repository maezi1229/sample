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
from quote_automation.scripts.update_ledger import append_ledger_entry, generate_quote_number
from quote_automation.template.fill_quote import (
    LIVE_FORMULA_ROUND_DIGITS,
    compute_customer_unit_price,
    excel_round,
    fill_quote_template,
)

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def build_output_filename(item_title: str) -> str:
    safe_title = INVALID_FILENAME_CHARS.sub("_", item_title).strip()
    return f"【見積書】{safe_title}.xlsx"


def effective_customer_unit_price(item, default_markup_percent: float) -> float:
    """品目の客先単価を決定する。単価直接指定があればそれを優先し、
    なければ品目別/全体デフォルトの上乗せ率から計算する。"""
    if item.customer_unit_price is not None:
        return item.customer_unit_price
    rate = item.markup_percent if item.markup_percent is not None else default_markup_percent
    if getattr(item, "live_formula", False):
        return excel_round(item.unit_price * (1 + rate / 100), LIVE_FORMULA_ROUND_DIGITS)
    return compute_customer_unit_price(item.unit_price, rate)


def run(confirmed_json_path: Path, template_path: Path, output_dir: Path, skip_ledger: bool = False) -> Path:
    confirmed = load_confirmed_quote(confirmed_json_path)

    print(f"客先名: {confirmed.customer_name}")
    print(f"担当者名: {confirmed.contact_name}")
    print(f"件名: {confirmed.item_title}")
    print(f"デフォルト上乗せ率: {confirmed.markup_percent}%（1000円単位で切り上げ、品目別に個別指定があればそちらを優先）")
    print(f"明細: {len(confirmed.items)}件")
    for item in confirmed.items:
        if item.customer_unit_price is not None:
            print(f"  - {item.name} 数量={item.qty} 仕入単価={item.unit_price:,.0f} "
                  f"客先単価={item.customer_unit_price:,.0f}（直接指定）")
        else:
            effective_rate = item.markup_percent if item.markup_percent is not None else confirmed.markup_percent
            note = "（Excel数式のまま。率を変えると自動再計算）" if item.live_formula else ""
            print(f"  - {item.name} 数量={item.qty} 仕入単価={item.unit_price:,.0f} 上乗せ率={effective_rate}%{note}")
    print(f"備考: {len(confirmed.remarks_lines)}行")
    if confirmed.receiving_place != "貴社車上渡し":
        print(f"受渡場所及び条件: {confirmed.receiving_place}")
    if confirmed.payment_terms != "従来通り":
        print(f"決済条件: {confirmed.payment_terms}")
    if confirmed.inspection_terms:
        print(f"検収条件: {confirmed.inspection_terms}")
    if confirmed.delivery_note:
        print(f"納期: {confirmed.delivery_note}")
    if confirmed.freight:
        f = confirmed.freight
        if f.customer_unit_price is not None:
            print(f"運賃: {f.name} 数量={f.qty} 仕入単価={f.unit_price:,.0f} 客先単価={f.customer_unit_price:,.0f}（直接指定）")
        else:
            rate = f.markup_percent if f.markup_percent is not None else confirmed.markup_percent
            print(f"運賃: {f.name} 数量={f.qty} 仕入単価={f.unit_price:,.0f} 上乗せ率={rate}%")
        print(f"運賃条件表記: {confirmed.freight_terms}")

    quote_number = confirmed.quote_number or generate_quote_number(config.LEDGER_PATH)
    print(f"見積№: {quote_number}")

    stamp_path = config.STAMP_IMAGE_PATH if confirmed.stamp else None
    if confirmed.stamp and not stamp_path.exists():
        print(f"捺印: 画像が見つからないため今回はスキップします（{stamp_path}）")
        stamp_path = None
    elif stamp_path:
        print("捺印: あり")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / build_output_filename(confirmed.item_title)

    def to_item_dict(i):
        return {
            "name": i.name,
            "qty": i.qty,
            "unit_price": i.unit_price,
            "markup_percent": i.markup_percent,
            "customer_unit_price": i.customer_unit_price,
            "live_formula": i.live_formula,
        }

    items = [to_item_dict(i) for i in confirmed.items]
    freight_dict = to_item_dict(confirmed.freight) if confirmed.freight else None

    stamp_kwargs = {"stamp_path": stamp_path}
    if stamp_path:
        if confirmed.stamp_cell:
            stamp_kwargs["stamp_cell"] = confirmed.stamp_cell
        if confirmed.stamp_offset_x_px is not None:
            stamp_kwargs["stamp_offset_x_px"] = confirmed.stamp_offset_x_px
        if confirmed.stamp_offset_y_px is not None:
            stamp_kwargs["stamp_offset_y_px"] = confirmed.stamp_offset_y_px
        if confirmed.stamp_size_px is not None:
            stamp_kwargs["stamp_size_px"] = confirmed.stamp_size_px

    result_path = fill_quote_template(
        template_path,
        output_path,
        customer_name=confirmed.customer_name,
        contact_name=confirmed.contact_name,
        item_title=confirmed.item_title,
        markup_percent=confirmed.markup_percent,
        items=items,
        remarks_lines=confirmed.remarks_lines,
        delivery_note=confirmed.delivery_note or None,
        freight=freight_dict,
        freight_terms=confirmed.freight_terms,
        receiving_place=confirmed.receiving_place,
        payment_terms=confirmed.payment_terms,
        inspection_terms=confirmed.inspection_terms or None,
        quote_number=quote_number,
        **stamp_kwargs,
    )
    print(f"転記済み見積書を出力しました: {result_path}")

    if not skip_ledger:
        all_priced_rows = list(confirmed.items) + ([confirmed.freight] if confirmed.freight else [])
        cost_amount = sum(i.qty * i.unit_price for i in all_priced_rows)
        sell_amount = sum(
            i.qty * effective_customer_unit_price(i, confirmed.markup_percent)
            for i in all_priced_rows
        )
        profit_amount = sell_amount - cost_amount
        # 品目ごとに上乗せ率が異なりうるため、集計表の利益率は「実際の利益÷仕入金額」で
        # 計算した実効レートを記録する（品目別レートを個別に指定していなければ、
        # 全品目共通のmarkup_percentと一致する）。
        blended_profit_rate = (profit_amount / cost_amount * 100) if cost_amount else 0.0
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
            profit_rate=round(blended_profit_rate, 2),
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
