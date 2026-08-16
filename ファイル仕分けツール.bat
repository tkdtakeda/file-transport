@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo  ファイル自動仕分けツール
echo ============================================================
echo.
echo ブラウザで操作画面を開きます。
echo ルールの設定、移動プレビュー、ファイル移動は、すべて画面から行えます。
echo.
echo このウィンドウはツールの動作に必要です。閉じないでください。
echo 終了するときは、画面右上の「…」から「ツールを終了」を選んでください。
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  py pdf_sorter_app.py
) else (
  python pdf_sorter_app.py
)

echo.
pause
