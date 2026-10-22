@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
python -c "import PIL" 2>nul || python -m pip install pillow
python -c "import fitz" 2>nul || python -m pip install pymupdf
python -c "import pypinyin" 2>nul || python -m pip install pypinyin
python -c "import opencc" 2>nul || python -m pip install opencc-python-reimplemented
python build_bom_data.py
pause
