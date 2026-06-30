@echo off
chcp 65001 >nul
title 征信报告OCR工具 - 打包程序

echo ============================================
echo   征信报告OCR工具 - Windows 打包程序
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装 Python 3.8+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python %version%

:: 安装依赖
echo.
echo [1/3] 安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [警告] 部分依赖安装失败，尝试继续...
)

:: 安装 PyInstaller
pip install pyinstaller
if %errorlevel% neq 0 (
    echo [错误] PyInstaller 安装失败
    pause
    exit /b 1
)
echo [OK] 依赖安装完成

:: 打包
echo.
echo [2/3] 开始打包（首次打包需下载EasyOCR模型，耗时约5-10分钟）...
pyinstaller --clean ^
    --windowed ^
    --onefile ^
    --name "征信报告OCR工具" ^
    --hidden-import PyQt5.sip ^
    --hidden-import easyocr ^
    --hidden-import easyocr.detection ^
    --hidden-import easyocr.recognition ^
    --hidden-import cv2 ^
    --hidden-import PIL._imaging ^
    --hidden-import openpyxl.cell._writer ^
    --hidden-import fitz ^
    --hidden-import numpy ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module pandas ^
    --exclude-module IPython ^
    --exclude-module notebook ^
    --add-data "README.md;." ^
    main.py

if %errorlevel% neq 0 (
    echo [错误] 打包失败
    pause
    exit /b 1
)
echo [OK] 打包完成

:: 清理临时文件
echo.
echo [3/3] 清理临时文件...
rmdir /s /q build 2>nul
del *.spec 2>nul

:: 复制输出
if exist "dist\征信报告OCR工具.exe" (
    echo.
    echo ============================================
    echo   ✅ 打包成功！
    echo   输出文件：dist\征信报告OCR工具.exe
    echo   大小：%~~z1 字节
    echo ============================================
    start dist
) else (
    echo [错误] 未找到打包后的文件
)

pause
