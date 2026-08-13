"""Step4: 転記・検証済みの見積書をPDF化する。

Step3の検証（利益上乗せ計算が正しいか）に通ったファイルだけをPDF化する。
検証がNGの場合はPDFを作らずに止まる。

使い方:
    python step4_export_pdf.py <Step2で作成したxlsx> <PDF出力先フォルダ>
"""
import argparse
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
from quote_automation import config
from quote_automation.template.export_pdf import export_to_pdf
from quote_automation.scripts.step3_verify_calculation import run as verify_calculation


def run(quote_path: Path, output_dir: Path) -> Path:
    print("=== Step3: 出力前の最終検証 ===")
    verify_calculation(quote_path)  # NGならVerificationErrorで例外送出、PDF化しない

    print()
    print("=== Step4: PDF化 ===")
    with tempfile.TemporaryDirectory() as tmp:
        profile_dir = Path(tmp) / "lo_profile"
        pdf_path = export_to_pdf(quote_path, output_dir, profile_dir)
    print(f"PDF出力先: {pdf_path}")
    return pdf_path


def main():
    parser = argparse.ArgumentParser(description="見積書xlsxを検証してからPDF化する")
    parser.add_argument("quote_path", type=Path)
    parser.add_argument("output_dir", type=Path, nargs="?", default=config.LOCAL_WORK_DIR / "output_pdf")
    args = parser.parse_args()
    run(args.quote_path, args.output_dir)


if __name__ == "__main__":
    main()
