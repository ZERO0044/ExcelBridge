@echo off
chcp 65001 >nul
echo ============================================
echo   ExcelBridge - Windows 打包脚本
echo   需要: Python 3.8 + pip
echo ============================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.8
    echo 下载: https://www.python.org/downloads/release/python-3810/
    pause
    exit /b 1
)

echo [1/4] 安装依赖...
pip install customtkinter openpyxl xlrd==1.2.0 Pillow pyinstaller

echo.
echo [2/4] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo [3/4] 开始打包 (约 2-5 分钟)...
pyinstaller ExcelBridge.spec

echo.
echo [4/4] 检查输出...
if exist "dist\ExcelBridge.exe" (
    echo.
    echo ============================================
    echo   ✅ 打包成功!
    echo   输出文件: dist\ExcelBridge.exe
    echo   大小:
    dir "dist\ExcelBridge.exe" | find "ExcelBridge.exe"
    echo ============================================
) else (
    echo [错误] 打包失败，请检查上方错误信息
)

echo.
echo 将 dist\ExcelBridge.exe 复制到任意 Windows 7+ 电脑即可运行。
pause
