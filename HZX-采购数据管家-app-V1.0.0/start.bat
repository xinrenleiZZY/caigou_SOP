@echo off
chcp 65001 >nul
title HZX-采购数据管家 V1
echo 正在启动 HZX-采购数据管家...
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo [错误] 启动失败，请确保已安装依赖：
    echo pip install -r requirements.txt
    pause
)
