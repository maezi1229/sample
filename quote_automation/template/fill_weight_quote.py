"""重量×キロ単価方式の見積書テンプレートへ、確定済みの見積データを転記する。

通常の見積（quote_automation/template/fill_quote.py、仕入単価×上乗せ率方式）とは
別の見積形式。鋼材の孔明け加工費など、「総重量(kg) × キロ単価」で金額を出す
案件向け（客先から受け取った実例をもとにテンプレート化した）。

テンプレート: quote_automation/assets/weight_based_quote_template.xlsx
- 明細は8枠（①〜⑧、19〜26行目）。1行 = 品名・式(数量)・キロ単価・総重量(kg)・
  金額(=キロ単価×総重量、テンプレート側の固定数式)・備考
- 見積有効期限は見積作成日(K3=TODAY())から+30日を自動計算する数式（E13）
- 見積№はK4に「見積№：〇〇」の形式で表示する（fill_quote.pyと同じ採番ルール）
- 捺印は宛先企業情報欄(K6:L11)付近に貼り付ける（fill_quote.pyと同じ画像・既定位置）
"""
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.units import pixels_to_EMU

ITEM_ROWS = tuple(range(19, 27))  # ①〜⑧、19〜26行目（1行1枠）
TOTAL_FORMULA_CELL = "D8"

QUOTE_NUMBER_CELL = "K4"
QUOTE_NUMBER_LABEL = "見積№"

DEFAULT_PAYMENT_TERMS = "従来通り"

OTHER_REMARKS_ROWS = (29, 30)  # 「その他備考」欄。最大2行

STAMP_DEFAULT_CELL = "K9"
STAMP_DEFAULT_OFFSET_X_PX = 40
STAMP_DEFAULT_OFFSET_Y_PX = 10
STAMP_DEFAULT_SIZE_PX = 80


class TooManyItemsError(Exception):
    pass


class TooManyOtherRemarksError(Exception):
    pass


class MissingRequiredFieldError(Exception):
    pass


def _require(value, label: str, missing: list) -> None:
    if value is None or (isinstance(value, str) and not value.strip()):
        missing.append(label)


def _add_stamp(
    ws,
    stamp_path: Path,
    cell: str = STAMP_DEFAULT_CELL,
    offset_x_px: int = STAMP_DEFAULT_OFFSET_X_PX,
    offset_y_px: int = STAMP_DEFAULT_OFFSET_Y_PX,
    size_px: int = STAMP_DEFAULT_SIZE_PX,
) -> None:
    img = XLImage(str(stamp_path))
    img.width = size_px
    img.height = size_px
    row, col = coordinate_to_tuple(cell)
    marker = AnchorMarker(
        col=col - 1, colOff=pixels_to_EMU(offset_x_px),
        row=row - 1, rowOff=pixels_to_EMU(offset_y_px),
    )
    img.anchor = OneCellAnchor(
        _from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(size_px), pixels_to_EMU(size_px)),
    )
    ws.add_image(img)


def fill_weight_quote_sheet(
    ws,
    *,
    customer_name: str,
    contact_name: str,
    item_title: str,
    items: list,
    receiving_place: str,
    payment_terms: str = DEFAULT_PAYMENT_TERMS,
    note: str = None,
    other_remarks_lines: list = None,
    quote_number: str = None,
    stamp_path: Path = None,
    stamp_cell: str = STAMP_DEFAULT_CELL,
    stamp_offset_x_px: int = STAMP_DEFAULT_OFFSET_X_PX,
    stamp_offset_y_px: int = STAMP_DEFAULT_OFFSET_Y_PX,
    stamp_size_px: int = STAMP_DEFAULT_SIZE_PX,
) -> None:
    """
    既に開いている1枚のシートに、確定済みの見積データを書き込む（ファイルI/Oなし）。

    items: [{"name": 品名, "qty": 式(数量、通常1), "kg_unit_price": キロ単価(円/kg),
             "total_weight_kg": 総重量(kg), "note": 備考(任意)}, ...] 最大8件まで。
        金額はテンプレート側の固定数式（キロ単価×総重量）で自動計算される。
    receiving_place: 受渡場所及び条件（案件ごとに指定。既定値は設けていない
        ―― 通常の見積書と違い、この形式は案件ごとに納入場所が大きく異なるため）。
    note: 見積条件欄の備考（例:「消費税別途」）。
    other_remarks_lines: 「その他備考」欄（最大2行、赤字太字の注意書き用）。
    """
    missing = []
    _require(customer_name, "客先名", missing)
    _require(contact_name, "担当者名", missing)
    _require(item_title, "件名（品名・物件名）", missing)
    _require(receiving_place, "受渡場所及び条件", missing)
    if missing:
        raise MissingRequiredFieldError(
            f"チャットでの確認が必要な項目が未指定です: {', '.join(missing)}。"
        )

    if len(items) > len(ITEM_ROWS):
        raise TooManyItemsError(
            f"テンプレートの明細枠は{len(ITEM_ROWS)}件までです（{len(items)}件指定されました）。"
        )
    other_remarks_lines = other_remarks_lines or []
    if len(other_remarks_lines) > len(OTHER_REMARKS_ROWS):
        raise TooManyOtherRemarksError(
            f"その他備考は{len(OTHER_REMARKS_ROWS)}行までです（{len(other_remarks_lines)}行指定されました）。"
        )

    ws["B2"] = customer_name
    ws["B3"] = contact_name
    ws["E10"] = f"：{item_title}"
    ws["E11"] = f"：{receiving_place}"
    ws["E12"] = f"：{payment_terms}"
    if note:
        ws["E14"] = f"：{note}"
    if quote_number:
        ws[QUOTE_NUMBER_CELL] = f"{QUOTE_NUMBER_LABEL}：{quote_number}"

    for row, item in zip(ITEM_ROWS, items):
        ws[f"C{row}"] = item["name"]
        ws[f"E{row}"] = item.get("qty", 1)
        ws[f"H{row}"] = item["kg_unit_price"]
        ws[f"I{row}"] = item["total_weight_kg"]
        if item.get("note"):
            ws[f"K{row}"] = item["note"]

    for row, line in zip(OTHER_REMARKS_ROWS, other_remarks_lines):
        ws[f"C{row}"] = line

    if stamp_path:
        _add_stamp(
            ws, stamp_path,
            cell=stamp_cell, offset_x_px=stamp_offset_x_px,
            offset_y_px=stamp_offset_y_px, size_px=stamp_size_px,
        )


def fill_weight_quote_template(template_path: Path, output_path: Path, **kwargs) -> Path:
    """テンプレートファイルを1件分の見積データで埋め、単独のxlsxとして保存する。
    kwargsはfill_weight_quote_sheet()と同じ。"""
    import shutil

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    fill_weight_quote_sheet(wb.active, **kwargs)

    wb.save(output_path)
    return output_path
