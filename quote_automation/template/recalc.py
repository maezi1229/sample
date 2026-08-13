"""LibreOffice Calc(headless)でxlsxの数式を再計算する。

openpyxlは数式を計算しない（保存時のキャッシュ値をそのまま読むだけ）ため、
実際にExcelで開いたときと同じ計算結果になっているかを確認するために使う。
Excelアプリそのものは起動しない・マクロも使わない。
"""
import shutil
import subprocess
from pathlib import Path


class RecalcError(Exception):
    pass


def recalculate_with_libreoffice(xlsx_path: Path, workdir: Path, profile_dir: Path) -> Path:
    xlsx_path = Path(xlsx_path)
    out_dir = Path(workdir)
    out_dir.mkdir(parents=True, exist_ok=True)

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to", "xlsx",
            "--outdir", str(out_dir),
            str(xlsx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise RecalcError(f"LibreOfficeでの再計算に失敗しました: {result.stdout}\n{result.stderr}")

    recalced = out_dir / xlsx_path.name
    if not recalced.exists():
        raise RecalcError(f"再計算後のファイルが見つかりません: {recalced}")
    return recalced
