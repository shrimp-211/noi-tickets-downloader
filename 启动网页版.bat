@echo off
chcp 936 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=gbk

where python >nul 2>nul
if errorlevel 1 (
    echo 未检测到 Python，请先安装 Python 3.8 及以上版本。
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

python web_app.py %*
if errorlevel 1 pause
