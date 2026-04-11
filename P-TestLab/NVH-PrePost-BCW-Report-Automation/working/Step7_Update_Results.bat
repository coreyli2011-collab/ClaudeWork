@echo off
title NVH Report — Step 7: Update Results
python "%~dp0update_results.py"
if errorlevel 1 (
    echo.
    echo Something went wrong. Please contact your administrator.
    pause
)
