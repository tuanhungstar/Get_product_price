@echo off
title Get Price Tool
echo [INFO] Initializing portable Python environment...
if not exist "python-embed\python.exe" (
    echo [ERROR] Python environment not found! Please run 'setup_env.bat' first.
    pause
    exit /b 1
)

echo [INFO] Launching Get Price application...
start "" "python-embed\pythonw.exe" Get_price.py
exit /b 0
