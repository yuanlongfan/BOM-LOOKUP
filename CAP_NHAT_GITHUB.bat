@echo off
chcp 65001 >nul
cd /d "%~dp0"
set REPO=https://github.com/yuanlongfan/bom-lookup.git

where git >nul 2>nul
if errorlevel 1 (
  echo Chua cai Git. Mo CMD va chay:  winget install Git.Git
  echo Cai xong dong cua so nay, bam dup lai file.
  pause
  exit /b 1
)

if not exist ".git" (
  git init -b main
  git remote add origin %REPO%
)
git remote set-url origin %REPO%

git add BOM_Lookup.html index.html README.md tools PICTURE
git commit -m "Cap nhat BOM %date% %time%"
git push -u origin main
if errorlevel 1 (
  echo.
  echo Day len that bai. Neu Git hoi dang nhap thi dang nhap tai khoan yuanlongfan.
  echo Neu bao "rejected" thi chay:  git push -u origin main --force
  pause
  exit /b 1
)
echo.
echo Xong. Trang web: https://yuanlongfan.github.io/bom-lookup/
pause
