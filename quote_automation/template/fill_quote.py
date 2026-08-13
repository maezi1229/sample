"""自社Excel見積書テンプレートへ、仕入先見積の確定データを転記する。

Excelアプリは起動しない。openpyxlでテンプレートファイルを直接コピー・編集する。

テンプレート（例: 「富士鋼業　見積」シート）の前提:
- 客先名・件名・支払条件・備考など、仕入先PDFから読み取れない情報は
  テンプレート側に既に入力済みであるという運用を前提とし、このモジュールは触らない。
- このモジュールが書き込むのは、仕入先見積に由来する「仕入数量・仕入単価」と、
  客先向けの「数量」のみ。単価・金額・利益はテンプレート側の数式
  （例: F19 = ROUND(N19*$M$17,-2)）で自動計算される。
- 明細行は 19-20行目=①, 21-22行目=②, 23-24行目=③ の3枠固定、
  25行目は運賃専用（このモジュールでは変更しない）。テンプレートの行構成が
  変わった場合はITEM_ROWSを合わせて調整すること。
"""
import shutil
from pathlib import Path

import openpyxl

ITEM_ROWS = (19, 21, 23)  # 明細行の先頭行（各ブロック2行分をマージしている）
MARGIN_CELL_ABS = "$M$17"  # 利益率セル（例: 1.09 = 9%上乗せ）


class TooManyItemsError(Exception):
    pass


class DestinationNotSetError(Exception):
    """客先名・宛先担当者名がテンプレートに未入力の場合に送出する。"""


def _check_destination_filled(ws) -> None:
    customer = ws["B2"].value
    contact = ws["B3"].value
    missing = []
    if not str(customer or "").strip():
        missing.append("B2（客先名）")
    if not str(contact or "").strip():
        missing.append("B3（宛先担当者名）")
    if missing:
        raise DestinationNotSetError(
            f"テンプレートに宛先が未入力です: {', '.join(missing)}。"
            "自動転記の前に、案件に合わせて客先名・担当者名を入力してください。"
        )


def _ensure_item_formulas(ws, row: int) -> None:
    """単価・金額・仕入金額・利益の数式が入っていることを保証する（①行目のパターンに合わせる）。"""
    ws[f"F{row}"] = f"=ROUND(N{row}*{MARGIN_CELL_ABS},-2)"
    ws[f"G{row}"] = f"=E{row}*F{row}"
    ws[f"O{row}"] = f"=M{row}*N{row}"
    ws[f"Q{row}"] = f"=G{row}-O{row}"


def fill_quote_template(template_path: Path, output_path: Path, items: list) -> Path:
    """
    items: [{"qty": 数量, "unit_price": 仕入単価, "name": 品名(任意)}, ...]
    最大3件まで（テンプレートの明細枠数に合わせる）。
    """
    if len(items) > len(ITEM_ROWS):
        raise TooManyItemsError(
            f"テンプレートの明細枠は{len(ITEM_ROWS)}件までです（{len(items)}件指定されました）。"
        )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    _check_destination_filled(ws)

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
