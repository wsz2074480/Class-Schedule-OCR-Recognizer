@echo off
cd /d "%~dp0"

call venv\Scripts\activate.bat

if "%~1"=="" (
    echo 请把课程表图片拖到这个 BAT 文件上。
    pause
    exit /b 1
)

python debug_pipeline.py "%~1"

pause
