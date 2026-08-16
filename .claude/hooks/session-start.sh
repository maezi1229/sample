#!/bin/bash
set -euo pipefail

# このリポジトリの見積書自動化ツール(quote_automation)が、Claude Code on the webの
# セッション開始時に毎回すぐ使えるように、必要なソフトを自動インストールする。
# 会社PC側には一切インストールしない（このスクリプトはクラウド側のセッションでのみ動く）。

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

apt-get update -qq
apt-get install -y -qq tesseract-ocr tesseract-ocr-jpn libreoffice-calc >/dev/null

pip install --quiet -r "$CLAUDE_PROJECT_DIR/quote_automation/requirements.txt"
