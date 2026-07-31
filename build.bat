@echo off
chcp 65001 >nul
echo ==========================================
echo   财务分析工具 - Windows 打包脚本
echo ==========================================
echo.

echo [1/4] 检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
python --version

echo.
echo [2/4] 安装/更新依赖...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [3/4] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist 财务分析工具.spec del 财务分析工具.spec

echo.
echo [4/4] 开始打包（单文件模式，无控制台窗口）...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "财务分析工具" ^
    --hidden-import PyQt5.sip ^
    --hidden-import openpyxl.cell._writer ^
    --collect-all pdfplumber ^
    --noconfirm ^
    main.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   打包完成！
echo   可执行文件: dist\财务分析工具.exe
echo ==========================================
echo.
echo 提示: 首次运行可能被 Windows Defender 拦截，选择「仍要运行」即可。
echo.
pause
