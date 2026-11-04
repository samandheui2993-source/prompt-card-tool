@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PS1="
for %%f in ("*.ps1") do set "PS1=%%~ff"
if not defined PS1 (
  echo [ERROR] No .ps1 gacha script found in this folder.
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
echo.
pause
