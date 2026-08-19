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

DELIVERY_CELL = "C15"
DELIVERY_LABEL = "納期" + "　" * 9  # 他の項目ラベル（受渡場所・決済条件等）と見た目の位置を合わせるための調整


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


def _set_delivery_note(ws, text: str) -> None:
    """納期はテンプレートに常設のフィールドではなく、指定があった案件だけ
    C15に追加する（未指定の見積は従来通り何も表示しない）。見た目を
    受渡場所・決済条件などの既存フィールドに合わせるため、C14のスタイルを
    そのままコピーしてから文言を書き込む。"""
    from copy import copy

    src = ws["C14"]
    dst = ws[DELIVERY_CELL]
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.fill = copy(src.fill)
    if not any(m.coord == "C15:D15" for m in ws.merged_cells.ranges):
        ws.merge_cells("C15:D15")
    dst.value = f"{DELIVERY_LABEL}：{text}"


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
    if delivery_note:
        _set_delivery_note(ws, delivery_note)
    _set_remarks(ws, remarks_lines or [])

    for row, item in zip(ITEM_ROWS, items):
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
                item_markup_percent = markup_percent
            customer_unit_price = compute_customer_unit_price(unit_price, item_markup_percent)
        _set_item_formulas(ws, row, customer_unit_price)
        # ②③の明細行（21-22, 23-24行目）はテンプレート側で初期状態は非表示になっている
        # （未使用時に空欄が印刷されないようにするため）。品目を入れた行は表示に切り替える。
        ws.row_dimensions[row].hidden = False
        ws.row_dimensions[row + 1].hidden = False
        # P列は印刷範囲(A1:J47)の外。この行の実効上乗せ率(%)をStep3の検算用に記録しておく
        # （品目ごとに率が異なりうるため、M17だけでは各行の期待値を再現できない）。
        # 単価を直接指定した品目はNoneのままにし、Step3では上乗せ率ベースの検算を行わない
        # （その場合はF列の値そのものが検算の基準になる）。
        ws[f"P{row}"] = item_markup_percent

    wb.save(output_path)
    return output_path
