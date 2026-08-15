@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo  PDF自動仕分け ルール編集画面
echo ============================================================
echo.
echo ブラウザでルール編集画面を開きます。
echo 終了するときは、このウィンドウで Ctrl + C を押してください。
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  py pdf_sorter_app.py server
) else (
  python pdf_sorter_app.py server
)

echo.
pause
