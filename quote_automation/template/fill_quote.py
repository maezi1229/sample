"""自社Excel見積書テンプレートへ、確定済みの見積データを転記する。

Excelアプリは起動しない。openpyxlでテンプレートファイルを直接コピー・編集する。

新仕様（チャット確認ベース）:
- 客先名・担当者名・件名（品名・物件名）・利益上乗せ率・備考は、
  仕入先PDFには書かれていない情報なので、Claudeがチャットで確認し、
  その回答をこのモジュールの引数として渡す。テンプレート側への事前入力は不要。
- 客先単価は「仕入単価 × (1 + 上乗せ率%)」を1000円単位で切り上げて作成する
  （例: F19 = ROUNDUP(N19*$M$17,-3)）。
- 見積有効期限は、この見積書の作成日（H3 = TODAY()）から+30日を自動計算する
  （C13に日付計算の数式を書き込むため、案件ごとに手入力しない）。
- 明細行は 19-20行目=①, 21-22行目=②, 23-24行目=③ の3枠固定、
  25行目は運賃専用（このモジュールでは変更しない）。
- 備考欄は34〜44行目を使う（29〜33行目は別用途の非表示行のため使わない。
  詳細はREADME参照）。
"""
import math
import shutil
from pathlib import Path

import openpyxl

ITEM_ROWS = (19, 21, 23)  # 明細行の先頭行（各ブロック2行分をマージしている）
MARGIN_CELL_ABS = "$M$17"  # 利益率セル（例: 1.09 = 9%上乗せ）
ROUND_DIGITS = -3  # 客先単価の丸め桁（-3 = 1000円単位）

REMARKS_FIRST_ROW = 34
REMARKS_LAST_ROW = 44  # 印刷範囲(A1:J47)に収まる範囲でここまで拡張可能

VALIDITY_CELL = "C13"
VALIDITY_FORMULA = '="見積有効期限         ："&TEXT(H3+30,"yyyy年m月d日")'


class TooManyItemsError(Exception):
    pass


class TooManyRemarksLinesError(Exception):
    pass


class MissingRequiredFieldError(Exception):
    """客先名・担当者名・件名・上乗せ率など、チャットで確認すべき項目が未指定の場合に送出する。"""


def excel_roundup(value: float, num_digits: int) -> float:
    """ExcelのROUNDUP関数（0から遠い方向への切り上げ）をPythonで再現する。
    テンプレート側の数式（F19等）と同じ計算をPython側でも独立に行うために使う
    （集計表への記帳、Step3の検算の両方で共通利用する）。
    """
    factor = 10 ** num_digits
    # 浮動小数点誤差で境界値がわずかに超過/不足すると切り上げ結果がずれるため、
    # 丸め桁より十分細かい精度でいったん丸めてから切り上げる（例: 109000.00000000001 -> 109000）。
    scaled = round(value * factor, 6)
    rounded = math.ceil(scaled) if scaled >= 0 else math.floor(scaled)
    return rounded / factor


def compute_customer_unit_price(unit_cost: float, markup_percent: float) -> float:
    return excel_roundup(unit_cost * (1 + markup_percent / 100), ROUND_DIGITS)


def _require(value, label: str, missing: list) -> None:
    if value is None or (isinstance(value, str) and not value.strip()):
        missing.append(label)


def _ensure_item_formulas(ws, row: int) -> None:
    """単価・金額・仕入金額・利益の数式が入っていることを保証する（①行目のパターンに合わせる）。"""
    ws[f"F{row}"] = f"=ROUNDUP(N{row}*{MARGIN_CELL_ABS},{ROUND_DIGITS})"
    ws[f"G{row}"] = f"=E{row}*F{row}"
    ws[f"O{row}"] = f"=M{row}*N{row}"
    ws[f"Q{row}"] = f"=G{row}-O{row}"


def _set_remarks(ws, lines: list) -> None:
    max_lines = REMARKS_LAST_ROW - REMARKS_FIRST_ROW + 1
    if len(lines) > max_lines:
        raise TooManyRemarksLinesError(
            f"備考が{len(lines)}行あり、テンプレートの備考欄（最大{max_lines}行）に収まりません。"
            "備考をまとめるか、テンプレートの拡張が必要です。"
        )
    for r in range(REMARKS_FIRST_ROW, REMARKS_LAST_ROW + 1):
        ws.cell(row=r, column=3).value = None
    for i, line in enumerate(lines):
        row = REMARKS_FIRST_ROW + i
        ws.cell(row=row, column=3).value = line
        ws.row_dimensions[row].hidden = False


def fill_quote_template(
    template_path: Path,
    output_path: Path,
    *,
    customer_name: str,
    contact_name: str,
    item_title: str,
    markup_percent: float,
    items: list,
    remarks_lines: list = None,
) -> Path:
    """
    items: [{"qty": 数量, "unit_price": 仕入単価, "name": 品名(任意)}, ...] 最大3件まで。
    markup_percent: 例えば9なら9%上乗せ（客先単価 = 仕入単価 * 1.09 を1000円単位で切り上げ）。
    remarks_lines: 仕入先見積の備考・注意事項をそのまま転記した行のリスト（チャットで事前確認済みのもの）。
    """
    missing = []
    _require(customer_name, "客先名", missing)
    _require(contact_name, "担当者名", missing)
    _require(item_title, "件名（品名・物件名）", missing)
    if markup_percent is None:
        missing.append("上乗せ率(%)")
    if missing:
        raise MissingRequiredFieldError(
            f"チャットでの確認が必要な項目が未指定です: {', '.join(missing)}。"
        )

    if len(items) > len(ITEM_ROWS):
        raise TooManyItemsError(
            f"テンプレートの明細枠は{len(ITEM_ROWS)}件までです（{len(items)}件指定されました）。"
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    ws["B2"] = customer_name
    ws["B3"] = contact_name
    ws["D10"] = item_title
    ws["M17"] = 1 + markup_percent / 100
    ws[VALIDITY_CELL] = VALIDITY_FORMULA
    _set_remarks(ws, remarks_lines or [])

    for row, item in zip(ITEM_ROWS, items):
        qty = item["qty"]
        unit_price = item["unit_price"]
        cost_qty = item.get("cost_qty", qty)

        ws[f"E{row}"] = qty
        ws[f"M{row}"] = cost_qty
        ws[f"N{row}"] = unit_price
        if item.get("name"):
            ws[f"C{row}"] = item["name"]
        _ensure_item_formulas(ws, row)

    wb.save(output_path)
    return output_path
