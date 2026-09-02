@echo off
chcp 65001 > nul
title 國中英語 L1 現在完成式與飲食文化 - 平板連線伺服器
echo ========================================================
echo  正在啟動 Lesson 1 教學伺服器，請確保電腦與平板連接同一 Wi-Fi...
echo ========================================================
python server.py
if errorlevel 1 (
    echo.
    echo 找不到 Python，嘗試使用 py 指令啟動...
    py server.py
)
pause
