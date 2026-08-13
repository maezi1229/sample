"""Step3: 利益上乗せ計算が正しいかを検証する。

Step2で転記した見積書(xlsx)をLibreOffice Calc(headless)で実際に再計算させ、
そこで得られた値が「仕入単価 × 利益率セル(M17)をROUND(-2)した値」と一致するかを
Pythonで独立に計算し直して突き合わせる。
一致しない場合は、テンプレートの数式が壊れている・利益率セルがずれている等の
異常を早期に検知できる。

使い方:
    python step3_verify_calculation.py <Step2で作成したxlsx>
"""
import argparse
import math
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation.template.fill_quote import ITEM_ROWS
from quote_automation.template.recalc import recalculate_with_libreoffice


class VerificationError(Exception):
    pass


def excel_round(value: float, num_digits: int) -> float:
    """ExcelのROUND関数（四捨五入、0から遠い方向）をPythonで再現する。"""
    factor = 10 ** num_digits
    scaled = value * factor
    rounded = math.floor(scaled + 0.5) if scaled >= 0 else math.ceil(scaled - 0.5)
    return rounded / factor


def run(quote_path: Path) -> bool:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        recalced_path = recalculate_with_libreoffice(quote_path, tmp / "recalced", tmp / "lo_profile")

        wb = openpyxl.load_workbook(recalced_path, data_only=True)
        ws = wb.active

        margin_rate = ws["M17"].value
        print(f"利益率設定 (M17): {margin_rate}（{(margin_rate - 1) * 100:.1f}%上乗せ）")
        print()

        all_ok = True
        total_customer = 0.0
        total_profit = 0.0

        for row in ITEM_ROWS:
            name = ws[f"C{row}"].value
            qty = ws[f"E{row}"].value
            unit_cost = ws[f"N{row}"].value
            customer_unit_price = ws[f"F{row}"].value
            customer_amount = ws[f"G{row}"].value
            cost_amount = ws[f"O{row}"].value
            profit = ws[f"Q{row}"].value

            if unit_cost is None:
                continue  # この行は未使用

            expected_unit_price = excel_round(unit_cost * margin_rate, -2)
            expected_amount = expected_unit_price * qty
            expected_profit = expected_amount - cost_amount

            ok = (
                customer_unit_price == expected_unit_price
                and customer_amount == expected_amount
                and abs((profit or 0) - expected_profit) < 1e-6
            )
            all_ok = all_ok and ok

            print(f"[{'OK' if ok else 'NG'}] {row}行目: {name}")
            print(f"      仕入単価={unit_cost:,.0f} × 数量{qty} → "
                  f"客先単価={customer_unit_price:,.0f}（期待値={expected_unit_price:,.0f}）")
            print(f"      客先金額={customer_amount:,.0f}（期待値={expected_amount:,.0f}） "
                  f"利益={profit:,.0f}（期待値={expected_profit:,.0f}）")

            total_customer += customer_amount
            total_profit += profit or 0

        grand_total = ws["G26"].value
        grand_profit = ws["Q26"].value
        totals_ok = (
            abs(grand_total - total_customer) < 1e-6
            and abs(grand_profit - total_profit) < 1e-6
        )
        all_ok = all_ok and totals_ok

        print()
        print(f"[{'OK' if totals_ok else 'NG'}] 合計金額: {grand_total:,.0f}円"
              f"（明細合計={total_customer:,.0f}円）")
        print(f"[{'OK' if totals_ok else 'NG'}] 合計利益: {grand_profit:,.0f}円"
              f"（明細合計={total_profit:,.0f}円）")
        print()
        print("検証結果: " + ("すべて一致（OK）" if all_ok else "不一致あり（NG） — テンプレートを確認してください"))

        if not all_ok:
            raise VerificationError("計算結果が期待値と一致しませんでした。")

        return all_ok


def main():
    parser = argparse.ArgumentParser(description="転記済み見積書の利益上乗せ計算を検証する")
    parser.add_argument("quote_path", type=Path)
    args = parser.parse_args()
    run(args.quote_path)


if __name__ == "__main__":
    main()
