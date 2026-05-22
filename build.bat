@echo off
chcp 65001 >nul
title 一键打包 轴模板处理工具

:: 进入脚本所在目录
cd /d "%~dp0"

echo [*] 检查 PyInstaller 是否安装...
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [!] PyInstaller 未安装，正在自动安装...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [X] 安装失败，请手动执行: pip install pyinstaller
        pause
        exit /b 1
    )
)

:: 可选：清理旧的打包残留
echo [*] 清理旧打包文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo [*] 开始打包，请稍候...
pyinstaller --onefile --windowed --name="抄轴工具" --add-data "HYWenHei-65W.ttf;." main.py

if %errorlevel% equ 0 (
    echo.
    echo ==========================================
    echo [√] 打包成功！
    echo exe 文件位于: %~dp0dist\抄轴工具.exe
    echo ==========================================
) else (
    echo [X] 打包失败，请检查错误信息。
)

pause