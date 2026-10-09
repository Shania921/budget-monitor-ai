@echo off
rem Budget Monitor AI - one-click entry (double-click me)
cd /d %~dp0
echo [check] python ...
python --version >nul 2>nul || (echo [error] Python not found in PATH. Install Python 3.9+ first. & pause & exit /b 1)
python -c "import openpyxl" >nul 2>nul || (echo [setup] installing openpyxl ... & python -m pip install openpyxl)
python src\run_all.py
pause
