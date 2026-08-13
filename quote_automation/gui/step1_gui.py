"""日常運用用：仕入先PDFをファイル選択ダイアログで選ぶだけでStep1を実行する。

コマンド入力に慣れていない人向けに、Windowsのバッチファイルからダブルクリックで
起動することを想定している（1_仕入先PDFを読み取る.bat 参照）。
"""
import os
import sys
import traceback
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation.scripts.step1_extract_supplier_quote import run as step1_run


def main():
    root = Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(
        title="① 仕入先から届いたPDF見積を選択してください",
        filetypes=[("PDFファイル", "*.pdf")],
    )
    if not pdf_path:
        return

    try:
        review_path = step1_run(Path(pdf_path))
    except Exception as e:
        messagebox.showerror("読み取りエラー", f"PDFの読み取りに失敗しました。\n\n{e}")
        traceback.print_exc()
        return

    messagebox.showinfo(
        "読み取り完了",
        "確認用シートを作成しました。\n\n"
        f"{review_path}\n\n"
        "これから開くExcelで、元のPDFと見比べながら「確定値」欄・明細（品名・数量・単価）を"
        "確認・修正して上書き保存してください。\n"
        "保存が終わったら「② 見積書を作成する」に進んでください。",
    )
    os.startfile(review_path)


if __name__ == "__main__":
    main()
