@echo off
cd /d "%~dp0"

call venv\Scripts\activate.bat

if "%~1"=="" (
    python parse_course_v5.py
) else (
    python parse_course_v5.py "%~1"
)

pause
