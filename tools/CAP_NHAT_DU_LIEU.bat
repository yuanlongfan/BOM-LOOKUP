@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
python -c "import PIL" 2>nul || python -m pip install pillow
python build_bom_data.py
pause
