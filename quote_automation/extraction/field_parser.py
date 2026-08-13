"""OCR/抽出テキストから見積書の項目を拾う（ベストエフォート）。

仕入先ごとに書式が異なり、OCRにも誤読が起こりうるため、
ここでの結果はあくまで「たたき台」。正しさの最終確認は人がレビューシートで行う前提。
"""
import re
from dataclasses import dataclass, field

DATE_RE = re.compile(r"(20\d{2})[/\-年](\d{1,2})[/\-月](\d{1,2})")
QUOTE_NO_RE = re.compile(r"No[.:\s]*([A-Za-z0-9\-]{4,})")
NUMBER_RE = re.compile(r"[\d,]+\.?\d*")

TABLE_HEADER_KEYWORDS = ("品名", "数量", "単価", "金額", "寸法", "員数", "重量", "規格")
NOISE_KEYWORDS = ("TEL", "FAX", "〒", "No.", "日付", "様", "御中", "件名", "見積書", "御依頼", "御用命")


def _normalize_ocr_numbers(text: str) -> str:
    """OCRがカンマの後に挿入する余分な空白を除去し、桁区切りが分断されるのを防ぐ。"""
    return re.sub(r",\s+(\d{3}\b)", r",\1", text)


@dataclass
class LineItem:
    raw_line: str
    name_guess: str = ""
    qty_guess: str = ""
    unit_price_guess: str = ""
    amount_guess: str = ""


@dataclass
class DraftQuote:
    quote_no: str = ""
    date: str = ""
    total_amount: str = ""
    line_items: list = field(default_factory=list)
    raw_text: str = ""


def _clean_amount(s: str) -> str:
    s = s.replace(",", "").replace(" ", "")
    # OCRが桁区切りのカンマをピリオドと誤読するケース（例: "383.000" -> "383000")
    if re.fullmatch(r"\d+\.\d{3}", s):
        s = s.replace(".", "")
    return s


_PAGE_FOOTER_RE = re.compile(r"^\(?\s*\d+\s*/\s*\d+\s*\)?$")


def parse_quote_no(text: str) -> str:
    m = QUOTE_NO_RE.search(text)
    return m.group(1) if m else ""


def parse_date(text: str) -> str:
    m = DATE_RE.search(text)
    if not m:
        return ""
    y, mo, d = m.groups()
    return f"{y}/{int(mo):02d}/{int(d):02d}"


def parse_total_amount(text: str) -> str:
    for line in text.splitlines():
        if "見積金額" in line or "御見積金額" in line or "【合" in line or "合　計" in line or "合計" in line:
            nums = re.findall(r"[\d,]{3,}", line)
            if nums:
                return _clean_amount(max(nums, key=len))
    return ""


def _find_item_table_window(lines: list) -> tuple:
    """明細表の見出し行〜合計行の間を特定する（電話番号・住所等の誤検出を防ぐ）。
    見出し行がOCRで読めなかった場合は文書全体をフォールバック対象にする。
    """
    header_idx = None
    for i, line in enumerate(lines):
        hits = sum(1 for kw in TABLE_HEADER_KEYWORDS if kw in line)
        if hits >= 2:
            header_idx = i
            break

    start_idx = header_idx + 1 if header_idx is not None else 0
    end_idx = len(lines)
    if header_idx is not None:
        for i in range(start_idx, len(lines)):
            if "合計" in lines[i] or "見積金額" in lines[i]:
                end_idx = i
                break
    return start_idx, end_idx, header_idx is not None


def parse_line_items(text: str) -> list:
    """明細表とおぼしき範囲で、行末の数値（数量/単価/金額想定）を拾う。
    見出し行が読めた場合は3つ、読めなかった場合は2つ以上の数値で判定する。
    """
    lines = text.splitlines()
    start_idx, end_idx, header_found = _find_item_table_window(lines)

    items = []
    for line in lines[start_idx:end_idx]:
        if any(kw in line for kw in NOISE_KEYWORDS):
            continue
        if _PAGE_FOOTER_RE.match(line.strip()):
            continue
        nums = NUMBER_RE.findall(line)
        nums = [n for n in nums if n.replace(",", "").replace(".", "")]
        qty = unit_price = amount = None
        if len(nums) >= 3:
            qty, unit_price, amount = nums[-3], nums[-2], nums[-1]
        elif len(nums) == 2:
            qty, unit_price, amount = "", nums[-2], nums[-1]

        if unit_price is not None:
            name_part = line
            for n in (qty, unit_price, amount):
                if n:
                    name_part = name_part.replace(n, "", 1)
            name_part = name_part.strip(" |:：")
            if name_part:
                items.append(LineItem(
                    raw_line=line.strip(),
                    name_guess=name_part,
                    qty_guess=_clean_amount(qty) if qty else "",
                    unit_price_guess=_clean_amount(unit_price),
                    amount_guess=_clean_amount(amount),
                ))
        elif header_found and line.strip() and not line.strip().startswith("【"):
            # 見出し行が特定できている場合のみ、数値の拾えない行（商品名が2行に分かれた
            # 場合の2行目など）も参考として残す。見出しが読めていない全文フォールバック時は
            # ノイズが増えるため対象外にする。
            items.append(LineItem(raw_line=line.strip(), name_guess=line.strip()))
    return items


def parse_draft(text: str) -> DraftQuote:
    text = _normalize_ocr_numbers(text)
    return DraftQuote(
        quote_no=parse_quote_no(text),
        date=parse_date(text),
        total_amount=parse_total_amount(text),
        line_items=parse_line_items(text),
        raw_text=text,
    )
