"""base_quote_template.xlsx の明細枠を3件からN件に拡張する、1回限りの構築スクリプト。

insert_rows()は値はずれた位置に移動するが、数式の行参照・結合セル範囲・
行の表示/非表示・印刷範囲は自動更新されないため、それらはすべてこのスクリプトで
明示的に再構築する。実行後は出力結果を目視確認してから
quote_automation/assets/base_quote_template.xlsx を置き換えること。
"""
import sys
from copy import copy
from pathlib import Path

import openpyxl

SRC = Path(sys.argv[1])
DST = Path(sys.argv[2])
N_ITEMS = int(sys.argv[3]) if len(sys.argv) > 3 else 17

wb = openpyxl.load_workbook(SRC)
ws = wb.active

COLS = list(range(2, 18))  # B..Q

# --- 元のブロック(19-20行目 / 21-22行目)の列スタイルを保存 ---
def capture_style(rows):
    style = {}
    for r_off, r in enumerate(rows):
        for c in COLS:
            cell = ws.cell(row=r, column=c)
            style[(r_off, c)] = {
                'font': copy(cell.font),
                'border': copy(cell.border),
                'fill': copy(cell.fill),
                'alignment': copy(cell.alignment),
                'number_format': cell.number_format,
            }
    return style

first_item_style = capture_style([19, 20])   # 品目①（C列に=D10のデフォルト式あり）
plain_item_style = capture_style([21, 22])   # 品目②（C列が空白のクリーンな型）
freight_style = capture_style([25, 26])      # 25行目=運賃行、26行目=合計行 のスタイルも保存

INSERT_AT = 25
INSERT_COUNT = 2 * (N_ITEMS - 3)  # 既存の3枠(①②③)より増える分
assert INSERT_COUNT >= 0, "N_ITEMS は3以上にしてください"

# openpyxlのinsert_rows()は挿入位置より下の結合セルとの相性に既知の不具合があるため
# （中身は動くのに結合範囲の座標だけ取り残され、挿入後の解除でKeyErrorになることがある）、
# 19行目以降の結合セルはinsert_rows()の前にいったんすべて解除し、挿入・再構築後に
# このスクリプトで作り直す。
for rng in [m.coord for m in list(ws.merged_cells.ranges) if m.min_row >= 19]:
    ws.unmerge_cells(rng)

if INSERT_COUNT > 0:
    ws.insert_rows(INSERT_AT, INSERT_COUNT)

last_item_row = 19 + 2 * N_ITEMS - 2  # 最後の品目ブロックの先頭行
freight_row = last_item_row + 2
total_row = freight_row + 1
spacer_row = total_row + 1
biko_label_row = spacer_row + 1
remarks_first = biko_label_row + 1
remarks_last = remarks_first + 10  # 11行分（従来と同じ容量）
print_last_row = remarks_last + 3

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"


def apply_style(r_top, style):
    for r_off in (0, 1):
        r = r_top + r_off
        for c in COLS:
            cell = ws.cell(row=r, column=c)
            s = style[(r_off, c)]
            cell.font = copy(s['font'])
            cell.border = copy(s['border'])
            cell.fill = copy(s['fill'])
            cell.alignment = copy(s['alignment'])
            cell.number_format = s['number_format']


def apply_style_single_row(r, style, style_row_offset):
    for c in COLS:
        cell = ws.cell(row=r, column=c)
        s = style[(style_row_offset, c)]
        cell.font = copy(s['font'])
        cell.border = copy(s['border'])
        cell.fill = copy(s['fill'])
        cell.alignment = copy(s['alignment'])
        cell.number_format = s['number_format']


def merge_item_block(r_top):
    """明細行（品目①〜）は2行1組のブロック。"""
    for rng in (f"B{r_top}:B{r_top+1}", f"C{r_top}:D{r_top+1}", f"E{r_top}:E{r_top+1}",
                f"F{r_top}:F{r_top+1}", f"G{r_top}:G{r_top+1}", f"H{r_top}:I{r_top+1}"):
        if not any(m.coord == rng for m in ws.merged_cells.ranges):
            ws.merge_cells(rng)


def merge_single_row(r):
    """運賃行・合計行は1行のみ（品目行のような2行ブロックではない）。"""
    for rng in (f"C{r}:D{r}", f"H{r}:I{r}"):
        if not any(m.coord == rng for m in ws.merged_cells.ranges):
            ws.merge_cells(rng)


# --- 品目行を19行目から N_ITEMS 件分、すべて明示的に作り直す ---
for i in range(N_ITEMS):
    r = 19 + 2 * i
    if i == 0:
        apply_style(r, first_item_style)
    else:
        apply_style(r, plain_item_style)
    merge_item_block(r)

    ws.cell(row=r, column=2, value=CIRCLED[i])  # B: 丸数字
    if i == 0:
        ws.cell(row=r, column=3, value="=D10")  # C: 品目①だけ件名を自動参照（従来仕様）
    else:
        ws.cell(row=r, column=3, value=None)
    ws.cell(row=r, column=5, value=None)   # E: 数量（fill_quote_templateが書く）
    ws.cell(row=r, column=6, value=f"=ROUNDUP(N{r}*$M$17,-3)")  # F: 一旦数式、後でPython値に上書きされる
    ws.cell(row=r, column=7, value=f"=E{r}*F{r}")  # G
    ws.cell(row=r, column=13, value=None)  # M: 仕入数量
    ws.cell(row=r, column=14, value=None)  # N: 仕入単価
    ws.cell(row=r, column=15, value=f"=M{r}*N{r}")  # O
    ws.cell(row=r, column=17, value=f"=G{r}-O{r}")  # Q
    ws.row_dimensions[r].hidden = (i != 0)       # 品目①だけ初期表示、他は使う時に表示
    ws.row_dimensions[r + 1].hidden = (i != 0)

# 品目①の数量・仕入数量だけ、従来のテンプレートに合わせて既定値1を残す
ws.cell(row=19, column=5, value=1)
ws.cell(row=19, column=13, value=1)

# --- 運賃行を作り直す（1行のみ。明細行のような2行ブロックではない） ---
apply_style_single_row(freight_row, freight_style, style_row_offset=0)
merge_single_row(freight_row)
ws.cell(row=freight_row, column=2, value=CIRCLED[N_ITEMS])  # 運賃は明細の次の番号
ws.cell(row=freight_row, column=3, value="運賃")
ws.cell(row=freight_row, column=5, value=1)
ws.cell(row=freight_row, column=6, value=None)
ws.cell(row=freight_row, column=7, value=f"=E{freight_row}*F{freight_row}")
ws.cell(row=freight_row, column=13, value=1)
ws.cell(row=freight_row, column=14, value=None)
ws.cell(row=freight_row, column=15, value=f"=M{freight_row}*N{freight_row}")
ws.cell(row=freight_row, column=16, value="運賃")
ws.cell(row=freight_row, column=17, value=f"=G{freight_row}-O{freight_row}")
ws.row_dimensions[freight_row].hidden = True

# --- 合計行（1行のみ） ---
apply_style_single_row(total_row, freight_style, style_row_offset=1)  # 26行目=合計行のスタイルを流用
merge_single_row(total_row)
ws.row_dimensions[total_row].height = 13.75
ws.cell(row=total_row, column=6, value="合計")
ws.cell(row=total_row, column=7, value=f"=SUM(G19:G{freight_row})")
ws.cell(row=total_row, column=17, value=f"=SUM(Q19:Q{freight_row})")

# --- 「備考」ラベル行 ---
ws.cell(row=biko_label_row, column=3, value="備考")
ws.row_dimensions[biko_label_row].height = 16.0

# --- 備考欄本体（11行、既存と同じ容量） ---
for i in range(11):
    r = remarks_first + i
    ws.row_dimensions[r].height = 16.0
    ws.row_dimensions[r].hidden = (i >= 4)  # 既存と同じく先頭4行だけ初期表示
    ws.cell(row=r, column=3, value=None)

# --- D8（総計金額）とC13（見積有効期限）の合計行参照を更新 ---
ws["D8"] = f"=G{total_row}"

# --- 印刷範囲を拡張 ---
ws.print_area = f"A1:J{print_last_row}"

print("last_item_row", last_item_row)
print("freight_row", freight_row)
print("total_row", total_row)
print("biko_label_row", biko_label_row)
print("remarks_first", remarks_first, "remarks_last", remarks_last)
print("print_last_row", print_last_row)

DST.parent.mkdir(parents=True, exist_ok=True)
wb.save(DST)
print("saved:", DST)
