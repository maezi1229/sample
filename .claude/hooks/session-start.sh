#!/bin/bash
set -euo pipefail

# このリポジトリの見積書自動化ツール(quote_automation)が、Claude Code on the webの
# セッション開始時に毎回すぐ使えるように、必要なソフトを自動インストールする。
# 会社PC側には一切インストールしない（このスクリプトはクラウド側のセッションでのみ動く）。

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

apt-get update -qq
apt-get install -y -qq tesseract-ocr tesseract-ocr-jpn libreoffice-calc fonts-noto-cjk >/dev/null

pip install --quiet -r "$CLAUDE_PROJECT_DIR/quote_automation/requirements.txt"

# 見積書テンプレートはWindows専用フォント「Meiryo UI」を指定しているが、この
# クラウド環境には入っていないため、見た目の近いNoto Sans CJK JPに差し替える
# （入れないとLibreOfficeがPDF変換時に任意のフォントへフォールバックし、
# 会社PCのExcel/Windowsで開いたときと文字の見た目が変わってしまう）。
cp "$CLAUDE_PROJECT_DIR/quote_automation/assets/fontconfig-local.conf" /etc/fonts/local.conf
fc-cache -f >/dev/null 2>&1
