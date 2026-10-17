@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==== DAY LEN GITHUB: bom-lookup (repo RIENG TU) ====
where git >nul 2>nul || (echo Chua cai Git. Tai tai https://git-scm.com/download/win roi chay lai. & pause & exit /b 1)
if not exist .git (
  git init -b main
)
git add -A
git commit -m "Tra cuu BOM quat tran" 2>nul
where gh >nul 2>nul
if %errorlevel%==0 (
  echo Dang nhap GitHub neu duoc hoi...
  gh auth status >nul 2>nul || gh auth login
  gh repo create bom-lookup --private --source . --remote origin --push
  if errorlevel 1 (git push -u origin main)
  echo.
  echo XONG. Moi dong nghiep: vao repo tren GitHub - Settings - Collaborators.
) else (
  echo Chua co GitHub CLI. Lam thu cong:
  echo  1. Mo https://github.com/new , dat ten bom-lookup, chon Private, bam Create.
  echo  2. Copy dong "git remote add origin ..." GitHub hien ra, dan vao cua so nay.
  echo  3. Chay: git push -u origin main
  start https://github.com/new
  cmd /k
)
pause
