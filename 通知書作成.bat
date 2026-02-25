@echo off
chcp 65001 > nul
echo ===================================================
echo 代理受領通知書 自動作成ツール
echo ===================================================
echo.

python "%~dp0generate_receipt.py"

pause
