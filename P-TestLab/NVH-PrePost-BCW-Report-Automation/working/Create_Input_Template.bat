@echo off
title NVH Report — Create Input Template
echo Creating Excel input template...
python "%~dp0create_input_template.py"
if errorlevel 1 (
    echo.
    echo Something went wrong. Please contact your administrator.
    pause
)
