"""チャットで確認済みの見積データ(JSON)を読み込む。

旧仕様（レビューシートxlsxを人がExcelで開いて確認）は廃止。
新仕様では、Claudeがチャット上で仕入先PDFの読み取り内容・客先名・担当者名・
上乗せ率(%)・備考をユーザーに確認し、その回答をこのJSON形式でまとめてから
Step2（テンプレート転記）に渡す。

JSONスキーマ（以下はダミーデータの例。実在の客先名・仕入先名・金額は書かない）:
{
  "customer_name": "サンプル工業株式会社 御中",
  "contact_name": "総務部　鈴木 様",
  "supplier_name": "テスト製鋼株式会社",
  "supplier_contact": "田中",
  "item_title": "ABC123_サンプル部品（１０×１０）",
  "markup_percent": 9,
  "delivery_note": "製品ご支給後、約2週間",
  "remarks_lines": ["溶接仕上げは前回製作時と同様になります。", ...],
  "items": [
    {"name": "ABC123_サンプル部品（１０×１０）", "qty": 3, "unit_price": 100000},
    {"name": "ABC456_サンプル加工費", "qty": 8, "unit_price": 30000, "markup_percent": 40},
    {"name": "ABC789_サンプル研磨費", "qty": 1, "unit_price": 150000, "customer_unit_price": 250000}
  ],
  "freight": {"unit_price": 1500, "customer_unit_price": 2000},
  "freight_terms": "別途運賃",
  "receiving_place": "客先工場（車上渡し）",
  "payment_terms": "検収月末締め翌月末振込",
  "inspection_terms": "搬入・試運転調整渡し"
}

supplier_name / supplier_contact は見積書テンプレートには転記しない
（自社見積書に仕入先名は出さない）。見積り集計表（分析用の記帳）にのみ使う。

items[].markup_percent は任意。品目ごとに上乗せ率を変えたい場合（例:
材料費は10%・加工費は40%、のように品目により率が異なる見積り）に指定する。
省略した品目はトップレベルのmarkup_percent（見積全体のデフォルト）を使う。

items[].customer_unit_price は任意。「〇〇円ちょうどにして」のように、
上乗せ率の計算ではなく客先単価を直接指定したい場合に使う。指定した品目は
markup_percent（品目別・全体デフォルトいずれも）より優先され、上乗せ率の
計算は行わずこの値をそのまま客先単価として使う。

items[].live_formula は任意（既定false）。trueにすると、その品目の単価を
Python側で計算した値ではなく、Excel数式「=ROUND(N×$M$17,0)」のまま残す。
上乗せ率セル(M17)を後からExcelで書き換えると単価が自動的に再計算される
（例:「20%で見積もったが、後で30%に変えるかもしれない」場合）。
1000円単位の切り上げは行わず、円単位の四捨五入になる。customer_unit_price
と併用不可。M17を共有するため、率の異なる複数品目がある見積りでは
1品目だけに使うことを想定している。

delivery_note（トップレベル）は任意。「製品ご支給後、約2週間」のように
納期を指定したい場合に使う。省略した場合は見積書に納期欄自体を表示しない
（従来通りの見た目になる）。

freight（トップレベル）は任意。運賃を明細（items、最大17件）とは別立てで
表示したい場合に使う。スキーマはitemsの要素と同じ（name省略時は「運賃」、
qty省略時は1）。指定しなければ従来通り運賃欄は表示しない。
freight_terms（トップレベル）は任意。見積条件欄の運賃表記（既定は
「運賃込み価格」）。freightで運賃を別立てにする場合は「別途運賃」等に変更する。

receiving_place / payment_terms（トップレベル）は任意。既定はそれぞれ
「貴社車上渡し」「従来通り」。客先ごとに決まった受渡場所・決済条件が
ある場合はそちらを指定する。
inspection_terms（トップレベル）は任意。検収条件を指定したい場合に使う。
省略した場合は見積書に検収条件欄自体を表示しない。

stamp（トップレベル、既定true）は任意。falseにすると、その見積りには
捺印を入れない。捺印画像はconfig.STAMP_IMAGE_PATH（Gitにはコミットしない
セッションローカルのファイル）が存在する場合のみ実際に貼り付けられる。
stamp_cell / stamp_offset_x_px / stamp_offset_y_px / stamp_size_px は
任意。捺印位置・サイズを既定（宛先企業情報欄付近）から変えたい場合に使う。
"""
import json
from dataclasses import dataclass, field
from pathlib import Path


class InvalidConfirmedQuoteError(Exception):
    """JSONの必須項目が欠けている場合に送出する。"""


@dataclass
class ConfirmedItem:
    name: str
    qty: float
    unit_price: float
    markup_percent: float = None  # 未指定ならConfirmedQuote.markup_percent(全体デフォルト)を使う
    customer_unit_price: float = None  # 指定時はmarkup_percentより優先し、この単価をそのまま使う
    live_formula: bool = False  # trueなら単価をExcel数式のまま残す（M17変更で自動再計算）


@dataclass
class ConfirmedQuote:
    customer_name: str
    contact_name: str
    item_title: str
    markup_percent: float
    items: list = field(default_factory=list)
    remarks_lines: list = field(default_factory=list)
    supplier_name: str = ""
    supplier_contact: str = ""
    delivery_note: str = ""
    freight: ConfirmedItem = None  # 運賃を明細と別立てにする場合のみ
    freight_terms: str = "運賃込み価格"
    receiving_place: str = "貴社車上渡し"
    payment_terms: str = "従来通り"
    inspection_terms: str = ""
    stamp: bool = True  # 既定で前島印を捺印する。不要な案件はfalseにする
    stamp_cell: str = None  # 捺印位置の上書き（既定はfill_quote.STAMP_DEFAULT_CELL）
    stamp_offset_x_px: int = None
    stamp_offset_y_px: int = None
    stamp_size_px: int = None


REQUIRED_TOP_LEVEL_KEYS = (
    "customer_name",
    "contact_name",
    "item_title",
    "markup_percent",
    "items",
)


def _parse_item(raw: dict, *, default_name: str = "", default_qty: float = None) -> ConfirmedItem:
    unit_price = raw.get("unit_price")
    qty = raw.get("qty", default_qty)
    if unit_price is None:
        raise InvalidConfirmedQuoteError(f"明細『{raw.get('name', default_name)}』の単価が未確定です。")
    if qty is None:
        raise InvalidConfirmedQuoteError(f"明細『{raw.get('name', default_name)}』の数量が未確定です。")
    item_markup = raw.get("markup_percent")
    item_price_override = raw.get("customer_unit_price")
    return ConfirmedItem(
        name=str(raw.get("name") or default_name).strip(),
        qty=float(qty),
        unit_price=float(unit_price),
        markup_percent=float(item_markup) if item_markup is not None else None,
        customer_unit_price=float(item_price_override) if item_price_override is not None else None,
        live_formula=bool(raw.get("live_formula", False)),
    )


def load_confirmed_quote(json_path: Path) -> ConfirmedQuote:
    json_path = Path(json_path)
    data = json.loads(json_path.read_text(encoding="utf-8"))

    missing = [k for k in REQUIRED_TOP_LEVEL_KEYS if not data.get(k) and data.get(k) != 0]
    if missing:
        raise InvalidConfirmedQuoteError(
            f"確定済み見積JSONに必須項目がありません: {', '.join(missing)}"
        )

    items = [_parse_item(raw) for raw in data["items"]]
    if not items:
        raise InvalidConfirmedQuoteError("明細（品名・数量・単価）が1件も確定されていません。")

    freight_raw = data.get("freight")
    freight = _parse_item(freight_raw, default_name="運賃", default_qty=1) if freight_raw else None

    return ConfirmedQuote(
        customer_name=str(data["customer_name"]).strip(),
        contact_name=str(data["contact_name"]).strip(),
        item_title=str(data["item_title"]).strip(),
        markup_percent=float(data["markup_percent"]),
        items=items,
        remarks_lines=list(data.get("remarks_lines") or []),
        supplier_name=str(data.get("supplier_name") or "").strip(),
        supplier_contact=str(data.get("supplier_contact") or "").strip(),
        delivery_note=str(data.get("delivery_note") or "").strip(),
        freight=freight,
        freight_terms=str(data.get("freight_terms") or "運賃込み価格").strip(),
        receiving_place=str(data.get("receiving_place") or "貴社車上渡し").strip(),
        payment_terms=str(data.get("payment_terms") or "従来通り").strip(),
        inspection_terms=str(data.get("inspection_terms") or "").strip(),
        stamp=bool(data.get("stamp", True)),
        stamp_cell=str(data["stamp_cell"]).strip() if data.get("stamp_cell") else None,
        stamp_offset_x_px=int(data["stamp_offset_x_px"]) if data.get("stamp_offset_x_px") is not None else None,
        stamp_offset_y_px=int(data["stamp_offset_y_px"]) if data.get("stamp_offset_y_px") is not None else None,
        stamp_size_px=int(data["stamp_size_px"]) if data.get("stamp_size_px") is not None else None,
    )
