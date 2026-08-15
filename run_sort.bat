@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo  PDF自動仕分け 実行
echo ============================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  py pdf_sorter_app.py sort
) else (
  python pdf_sorter_app.py sort
)

echo.
echo 処理が終わりました。
pause
