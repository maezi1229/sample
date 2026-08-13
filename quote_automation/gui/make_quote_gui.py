"""日常運用用：レビューシートとテンプレートを選ぶだけでStep2〜4を一気に実行する。

内部では
  Step2: レビューシートの確定値をテンプレートへ転記
  Step3: 利益上乗せ計算が正しいか自動検証（NGならPDF化しない）
  Step4: PDF化
を順番に行う。ダブルクリックで起動する想定（2_見積書を作成する.bat 参照）。
"""
import os
import sys
import traceback
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation import config
from quote_automation.extraction.confirmed_reader import ReviewNotConfirmedError
from quote_automation.scripts.step2_fill_template import run as step2_run
from quote_automation.scripts.step3_verify_calculation import VerificationError
from quote_automation.scripts.step4_export_pdf import run as step4_run
from quote_automation.template.fill_quote import DestinationNotSetError, TooManyItemsError


def main():
    root = Tk()
    root.withdraw()

    review_path = filedialog.askopenfilename(
        title="② 確認・修正済みのレビューシート(xlsx)を選択してください",
        filetypes=[("Excelファイル", "*.xlsx")],
    )
    if not review_path:
        return

    template_path = filedialog.askopenfilename(
        title="この案件用の見積書（客先名・件名など入力済みのテンプレート）を選択してください",
        filetypes=[("Excelファイル", "*.xlsx")],
    )
    if not template_path:
        return

    review_path = Path(review_path)
    output_dir = config.LOCAL_WORK_DIR / "output"
    output_xlsx = output_dir / (review_path.stem.replace("_review", "") + "_quote.xlsx")

    try:
        step2_run(review_path, Path(template_path), output_xlsx)
        pdf_path = step4_run(output_xlsx, config.LOCAL_WORK_DIR / "output_pdf")
    except ReviewNotConfirmedError as e:
        messagebox.showerror("確認が未完了です", f"{e}\n\nレビューシートを開いて確定値を入力してから、もう一度実行してください。")
        return
    except DestinationNotSetError as e:
        messagebox.showerror("宛先が未入力です", str(e))
        return
    except TooManyItemsError as e:
        messagebox.showerror("明細が多すぎます", str(e))
        return
    except VerificationError as e:
        messagebox.showerror(
            "計算結果が一致しません",
            "利益上乗せの計算結果が期待値と一致しなかったため、PDFは作成していません。\n"
            "テンプレートの数式や利益率(M17)が変更されていないか確認してください。\n\n"
            f"{e}",
        )
        return
    except Exception as e:
        messagebox.showerror("エラー", f"見積書の作成に失敗しました。\n\n{e}")
        traceback.print_exc()
        return

    messagebox.showinfo(
        "作成完了",
        f"見積書PDFを作成しました。\n\n{pdf_path}\n\n内容を確認してから、客先へメールで送付してください。",
    )
    os.startfile(pdf_path)


if __name__ == "__main__":
    main()
