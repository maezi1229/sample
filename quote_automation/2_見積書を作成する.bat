@echo off
cd /d "%~dp0"
py -3 gui\make_quote_gui.py
if errorlevel 1 (
    echo エラーが発生しました。
    pause
)
