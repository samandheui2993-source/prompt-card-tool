@echo off
rem ---- One-click launcher for the web gacha (opens in default browser) ----
cd /d "%~dp0"
set "HTML="
for %%f in ("*.html") do if not defined HTML set "HTML=%%~ff"
if not defined HTML (
  echo [ERROR] No .html file found in this folder.
  pause
  exit /b 1
)
start "" "%HTML%"
exit /b 0
