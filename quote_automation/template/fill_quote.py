"""自社Excel見積書テンプレートへ、確定済みの見積データを転記する。

Excelアプリは起動しない。openpyxlでテンプレートファイルを直接コピー・編集する。

新仕様（チャット確認ベース）:
- 客先名・担当者名・件名（品名・物件名）・利益上乗せ率・備考は、
  仕入先PDFには書かれていない情報なので、Claudeがチャットで確認し、
  その回答をこのモジュールの引数として渡す。テンプレート側への事前入力は不要。
- 客先単価は「仕入単価 × (1 + 上乗せ率%)」を1000円単位で切り上げて作成する。
  上乗せ率は品目ごとに個別指定できる（例: 材料費10%・加工費40%、のように
  品目により率が異なるケース）。品目に上乗せ率が指定されなければ、
  見積全体のデフォルト上乗せ率(markup_percent)を使う。
  客先単価はPython側で計算し、F列にはその計算結果を値として書き込む
  （品目ごとに率が異なりうるため、$M$17を参照する数式では表現できない。
  M17には参考としてデフォルト上乗せ率のみを表示用に入れる）。
- 「〇〇円ちょうどにして」のように上乗せ率ではなく客先単価そのものを
  指定したい品目は、item["customer_unit_price"]で直接指定できる
  （markup_percentより優先。上乗せ率の計算は行わずそのまま使う）。
- 見積有効期限は、この見積書の作成日（H3 = TODAY()）から+30日を自動計算する
  （C13に日付計算の数式を書き込むため、案件ごとに手入力しない）。
- 明細行は19行目から2行1組×17枠（①〜⑰、19-20,21-22,...,51-52行目）、
  53行目は運賃専用（このモジュールでは通常変更しない）。この枠数は
  quote_automation/scripts/_rebuild_template_for_n_items.py で
  base_quote_template.xlsxを再構築して作った数。さらに増やす場合は
  そのスクリプトのN_ITEMSを変えて再実行し、このファイルの行定数も合わせて直す。
- 備考欄は57〜67行目を使う（詳細はREADME参照）。
"""
import math
import shutil
from pathlib import Path

import openpyxl

ITEM_ROWS = tuple(range(19, 52, 2))  # 明細行の先頭行（各ブロック2行分をマージしている）。19,21,...,51 の17枠
FREIGHT_ROW = 53  # 運賃専用行。テンプレート側でC53="運賃"・数量1が既定済み
TOTAL_ROW = 54  # 合計行（G54=SUM(G19:G53), Q54=SUM(Q19:Q53)）
MARGIN_CELL_ABS = "$M$17"  # 利益率セル（例: 1.09 = 9%上乗せ）
ROUND_DIGITS = -3  # 客先単価の丸め桁（-3 = 1000円単位）

FREIGHT_TERMS_CELL = "C14"
DEFAULT_FREIGHT_TERMS = "運賃込み価格"  # 「送料は別途」等、案件に応じて上書きできる

REMARKS_FIRST_ROW = 57
REMARKS_LAST_ROW = 67  # 印刷範囲内に収まる範囲でここまで拡張可能

VALIDITY_CELL = "C13"
VALIDITY_FORMULA = '="見積有効期限         ："&TEXT(H3+30,"yyyy年m月d日")'

RECEIVING_CELL = "C11"
RECEIVING_LABEL = "受渡場所　及び条件　"
DEFAULT_RECEIVING_PLACE = "貴社車上渡し"

PAYMENT_TERMS_CELL = "C12"
PAYMENT_TERMS_LABEL = "決済条件               "
DEFAULT_PAYMENT_TERMS = "従来通り"

DELIVERY_CELL = "C15"
DELIVERY_LABEL = "納期" + "　" * 9  # 他の項目ラベル（受渡場所・決済条件等）と見た目の位置を合わせるための調整

INSPECTION_CELL = "C16"
INSPECTION_LABEL = "検収条件" + "　" * 7


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


def _set_item_formulas(ws, row: int, customer_unit_price: float) -> None:
    """単価は品目ごとの上乗せ率で計算済みの値をそのまま入れ、金額・仕入金額・利益は数式で持たせる
    （①行目のパターンに合わせる）。"""
    ws[f"F{row}"] = customer_unit_price
    ws[f"G{row}"] = f"=E{row}*F{row}"
    ws[f"O{row}"] = f"=M{row}*N{row}"
    ws[f"Q{row}"] = f"=G{row}-O{row}"


def _apply_priced_row(ws, row: int, item: dict, default_markup_percent: float) -> None:
    """1行分（明細行または運賃行）に、数量・仕入単価・客先単価を書き込む共通処理。"""
    qty = item["qty"]
    unit_price = item["unit_price"]
    cost_qty = item.get("cost_qty", qty)
    price_override = item.get("customer_unit_price")

    ws[f"E{row}"] = qty
    ws[f"M{row}"] = cost_qty
    ws[f"N{row}"] = unit_price
    if item.get("name"):
        ws[f"C{row}"] = item["name"]

    if price_override is not None:
        customer_unit_price = price_override
        item_markup_percent = None  # 手動指定単価。Step3側では上乗せ率チェックを行わない
    else:
        item_markup_percent = item.get("markup_percent")
        if item_markup_percent is None:
            item_markup_percent = default_markup_percent
        customer_unit_price = compute_customer_unit_price(unit_price, item_markup_percent)
    _set_item_formulas(ws, row, customer_unit_price)
    # P列は印刷範囲(A1:J47)の外。この行の実効上乗せ率(%)をStep3の検算用に記録しておく
    # （品目ごとに率が異なりうるため、M17だけでは各行の期待値を再現できない）。
    # 単価を直接指定した品目はNoneのままにし、Step3では上乗せ率ベースの検算を行わない
    # （その場合はF列の値そのものが検算の基準になる）。
    ws[f"P{row}"] = item_markup_percent


def _set_optional_terms_row(ws, cell_coord: str, label: str, text: str) -> None:
    """納期・検収条件など、テンプレートに常設ではなく指定があった案件だけ追加する
    行に書き込む共通処理（未指定の見積は従来通り何も表示しない）。見た目を
    受渡場所・決済条件などの既存フィールドに合わせるため、C14のスタイルを
    そのままコピーしてから文言を書き込む。"""
    from copy import copy

    src = ws["C14"]
    dst = ws[cell_coord]
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.fill = copy(src.fill)
    row = dst.row
    merge_rng = f"C{row}:D{row}"
    if not any(m.coord == merge_rng for m in ws.merged_cells.ranges):
        ws.merge_cells(merge_rng)
    ws.row_dimensions[row].height = ws.row_dimensions[14].height
    dst.value = f"{label}：{text}"


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
    delivery_note: str = None,
    freight: dict = None,
    freight_terms: str = DEFAULT_FREIGHT_TERMS,
    receiving_place: str = DEFAULT_RECEIVING_PLACE,
    payment_terms: str = DEFAULT_PAYMENT_TERMS,
    inspection_terms: str = None,
) -> Path:
    """
    items: [{"qty": 数量, "unit_price": 仕入単価, "name": 品名(任意),
             "markup_percent": 品目別の上乗せ率(任意、省略時はmarkup_percentを使う),
             "customer_unit_price": 客先単価を直接指定(任意、指定時はmarkup_percentより優先)}, ...]
           最大3件まで。
    markup_percent: 見積全体のデフォルト上乗せ率。例えば9なら9%上乗せ
        （客先単価 = 仕入単価 * 1.09 を1000円単位で切り上げ）。品目ごとに
        個別の上乗せ率が指定されていれば、その品目はそちらを優先する。
    remarks_lines: 仕入先見積の備考・注意事項をそのまま転記した行のリスト（チャットで事前確認済みのもの）。
    delivery_note: 納期（例:「製品ご支給後、約2週間」）。指定があった案件だけ表示する。
    freight: 運賃を明細と別立てにしたい場合に指定する
        {"qty": 数量(省略時1), "unit_price": 仕入単価, "name": 品目名(省略時「運賃」),
         "markup_percent": 任意, "customer_unit_price": 任意}。
        テンプレート専用の運賃行（FREIGHT_ROW）に書き込む（明細枠とは別）。
    freight_terms: 見積条件欄の運賃表記（既定は「運賃込み価格」）。運賃を明細で
        別立てにする場合は「別途運賃」等に変更する。
    receiving_place: 受渡場所及び条件（既定は「貴社車上渡し」）。客先ごとに
        決まった条件がある場合はそちらに置き換える。
    payment_terms: 決済条件（既定は「従来通り」）。客先ごとに決まった条件が
        ある場合はそちらに置き換える。
    inspection_terms: 検収条件。指定があった案件だけ表示する（既定では欄自体を表示しない）。
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
    ws[RECEIVING_CELL] = f"{RECEIVING_LABEL}：{receiving_place}"
    ws[PAYMENT_TERMS_CELL] = f"{PAYMENT_TERMS_LABEL}：{payment_terms}"
    ws[FREIGHT_TERMS_CELL] = f"備考：{freight_terms}"
    if delivery_note:
        _set_optional_terms_row(ws, DELIVERY_CELL, DELIVERY_LABEL, delivery_note)
    if inspection_terms:
        _set_optional_terms_row(ws, INSPECTION_CELL, INSPECTION_LABEL, inspection_terms)
    _set_remarks(ws, remarks_lines or [])

    for row, item in zip(ITEM_ROWS, items):
        _apply_priced_row(ws, row, item, markup_percent)
        # ①以外の明細行はテンプレート側で初期状態は非表示になっている
        # （未使用時に空欄が印刷されないようにするため）。品目を入れた行は表示に切り替える。
        ws.row_dimensions[row].hidden = False
        ws.row_dimensions[row + 1].hidden = False

    if freight:
        freight_item = dict(freight)
        freight_item.setdefault("qty", 1)
        _apply_priced_row(ws, FREIGHT_ROW, freight_item, markup_percent)
        ws.row_dimensions[FREIGHT_ROW].hidden = False

    wb.save(output_path)
    return output_path
