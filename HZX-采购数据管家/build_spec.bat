@echo off
chcp 65001 >nul
title HZX-采购数据管家 - 打包工具 (Spec模式)

echo ╔══════════════════════════════════════════╗
echo ║    HZX-采购数据管家  V1.0.0              ║
echo ║    开发者：IT-钟                          ║
echo ║    电话：18072740843                     ║
echo ╚══════════════════════════════════════════╝
echo.

:: 清理旧构建
echo [1/3] 清理旧构建...
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "build" rmdir /s /q "build" >nul 2>&1

:: 打包
echo [2/3] 打包中...
pyinstaller "HZX-采购数据管家.spec" --noconfirm

echo.
if exist "dist\HZX-采购数据管家.exe" (
    echo [3/3] 打包成功！
    echo.
    echo ╔══════════════════════════════════════════╗
    echo ║  输出: dist\HZX-采购数据管家.exe          ║
    for %%f in ("dist\HZX-采购数据管家.exe") do (
        set size=%%~zf
        set /a mb = size / 1048576
        echo ║  大小: %%mb MB
    )
    echo ╚══════════════════════════════════════════╝
) else (
    echo [错误] 打包失败！
)

pause
