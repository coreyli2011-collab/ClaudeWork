@echo off
title NVH Report — Step 2: Generate Reports
python "%~dp0generate_reports.py"
if errorlevel 1 (
    echo.
    echo Something went wrong. Please contact your administrator.
    pause
)
