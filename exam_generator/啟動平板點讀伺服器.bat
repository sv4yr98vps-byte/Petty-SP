@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 考卷點讀與教學簡報 本地伺服器
echo 正在啟動平板點讀伺服器...
"啟動平板點讀伺服器.exe"
pause
