@echo off
chcp 65001 >nul
title HZX-采购数据管家 - 开发者打包脚本

echo ╔══════════════════════════════════════════╗
echo ║    HZX-采购数据管家  V1.0.0              ║
echo ║    开发者：IT-钟                          ║
echo ║    电话：18072740843                     ║
echo ╚══════════════════════════════════════════╝
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python环境！
    pause
    exit /b 1
)

:: 安装依赖
echo [1/3] 安装依赖包...
pip install -r requirements.txt
echo.

:: 检查关键依赖
python -c "import PyQt5" 2>nul || ( echo [错误] PyQt5安装失败 & pause & exit /b 1 )
python -c "import qfluentwidgets" 2>nul || ( echo [错误] qfluentwidgets安装失败 & pause & exit /b 1 )
python -c "import openpyxl" 2>nul || ( echo [错误] openpyxl安装失败 & pause & exit /b 1 )
echo [OK] 依赖检查通过

:: 清理
echo [2/3] 清理旧构建...
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "build" rmdir /s /q "build" >nul 2>&1
if exist "HZX-采购数据管家.spec" del "HZX-采购数据管家.spec" >nul 2>&1

:: 打包
echo [3/3] 开始打包...
pyinstaller --noconfirm --onefile --windowed ^
    --name "HZX-采购数据管家" ^
    --hidden-import win32com ^
    --hidden-import openpyxl ^
    --hidden-import win32com.client ^
    --collect-all qfluentwidgets ^
    --add-data "config.py;." ^
    --add-data "main_processing.py;." ^
    --add-data "pages;pages" ^
    app.py

echo.
if exist "dist\HZX-采购数据管家.exe" (
    echo ╔══════════════════════════════════════════╗
    echo ║  打包成功！                              ║
    echo ║  输出: dist\HZX-采购数据管家.exe          ║
    echo ║  大小: 
    for %%f in ("dist\HZX-采购数据管家.exe") do echo ║  %%~zf 字节
    echo ╚══════════════════════════════════════════╝
) else (
    echo [错误] 打包失败，请检查错误日志
)

pause
