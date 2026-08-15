@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo  PDF自動仕分け 移動プレビュー（ファイルは移動しません）
echo ============================================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  py pdf_sorter_app.py preview
) else (
  python pdf_sorter_app.py preview
)

echo.
pause
