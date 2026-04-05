@echo off
chcp 65001 > nul
echo 正在启动发票归集软件...
echo.

REM 检查Python是否已安装
python --version > nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python。请先安装Python 3.8或更高版本。
    pause
    exit /b 1
)

REM 检查依赖包
echo 检查依赖包...
python -c "import sys; sys.exit(0)" > nul 2>&1
if errorlevel 1 (
    echo 错误: Python环境异常。
    pause
    exit /b 1
)

echo 安装依赖包...
pip install PyQt6 PyQt6-WebEngine pandas openpyxl PyMuPDF

echo.
echo 启动发票归集软件...
python app.py

pause