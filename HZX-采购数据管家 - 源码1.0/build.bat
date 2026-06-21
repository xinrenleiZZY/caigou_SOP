@echo off
chcp 65001 >nul
title HZX-采购数据管家 - 打包工具

echo ========================================
echo  HZX-采购数据管家 V1.0.0
echo  正在打包为exe...
echo ========================================
echo.

:: 安装依赖
pip install -r requirements.txt

:: 清理旧的构建
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"

:: PyInstaller 打包
pyinstaller --noconfirm --onefile --windowed ^
    --name "HZX-采购数据管家" ^
    --icon NONE ^
    --add-data "SOP_docs;SOP_docs" ^
    --hidden-import win32com ^
    --hidden-import openpyxl ^
    --collect-all qfluentwidgets ^
    app.py

echo.
echo ========================================
echo  打包完成！
echo  输出文件: dist\HZX-采购数据管家.exe
echo ========================================
pause
