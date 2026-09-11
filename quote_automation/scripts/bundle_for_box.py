"""見積り一式を、会社PCのBox同期フォルダに置きやすい形のzipにまとめる。

会社PC側のフォルダ構成（例）:
    C:\\Box\\1.厚板部_名古屋\\1.個人フォルダ\\前島\\1.営業\\顧客\\<客先名>\\<案件名>\\

Claude Code on the webのクラウドセッションからは会社PCのCドライブに直接
書き込めないため、案件ごとに「<客先名>/<案件名>/」というフォルダ構成の
zipを作り、それをダウンロードして解凍・配置してもらう運用にしている。
フォルダ名・保存要否は毎回チャットで確認してから実行すること
（README「Boxフォルダへの保存」参照）。

同梱するファイルには、見積の元になった仕入先見積PDFや作成した見積Excelだけでなく、
客先から依頼時に添付された図面・仕様書等の資料もあれば含める（ファイルリストに
そのパスを追加するだけでよい。フォルダ分けはせず案件フォルダ直下にまとめて置く）。

文字化け対策について:
    WindowsのExplorer標準の「すべて展開」は、zip内のファイル名がUTF-8で
    格納されていても（UTF-8フラグを無視して）システムのANSIコードページ
    （日本語Windowsでは Shift-JIS/cp932）で decode してしまうことがあり、
    その場合ファイル名が文字化けする。これを避けるため、このスクリプトは
    ファイル名をUTF-8ではなくcp932でエンコードし、UTF-8フラグを立てずに
    zipへ格納する（`_Cp932ZipInfo`）。

使い方:
    python bundle_for_box.py <客先名> <案件名> <zip出力先> <同梱するファイル...>
"""
import argparse
import time
import zipfile
from pathlib import Path


class _Cp932ZipInfo(zipfile.ZipInfo):
    """ファイル名をcp932で格納し、UTF-8フラグを立てない ZipInfo。

    Windows Explorer標準展開機能での日本語ファイル名の文字化けを防ぐため。
    """

    def _encodeFilenameFlags(self):
        try:
            return self.filename.encode("ascii"), self.flag_bits
        except UnicodeEncodeError:
            return self.filename.encode("cp932"), self.flag_bits


def _add_file(zf: zipfile.ZipFile, path: Path, arcname: str) -> None:
    date_time = time.localtime(path.stat().st_mtime)[:6]
    zinfo = _Cp932ZipInfo(arcname, date_time)
    zinfo.compress_type = zipfile.ZIP_DEFLATED
    zinfo.external_attr = 0o600 << 16
    zf.writestr(zinfo, path.read_bytes())


def bundle_for_box(customer_name: str, job_name: str, zip_path: Path, files: list) -> Path:
    zip_path = Path(zip_path)
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    arc_root = f"{customer_name}/{job_name}"

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            f = Path(f)
            if not f.exists():
                raise FileNotFoundError(f"ファイルが見つかりません: {f}")
            _add_file(zf, f, f"{arc_root}/{f.name}")

    return zip_path


def main():
    parser = argparse.ArgumentParser(description="見積り一式をBox配置用のzipにまとめる")
    parser.add_argument("customer_name")
    parser.add_argument("job_name")
    parser.add_argument("zip_path", type=Path)
    parser.add_argument("files", nargs="+", type=Path)
    args = parser.parse_args()

    result = bundle_for_box(args.customer_name, args.job_name, args.zip_path, args.files)
    print(f"作成しました: {result}")


if __name__ == "__main__":
    main()
