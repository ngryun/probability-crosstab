@echo off
chcp 65001 >nul
set "APP_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" "%APP_DIR%crosstab_app.py" "%APP_DIR%data.xlsx"
  goto end
)

py -3 "%APP_DIR%crosstab_app.py" "%APP_DIR%data.xlsx"
if errorlevel 1 (
  python "%APP_DIR%crosstab_app.py" "%APP_DIR%data.xlsx"
)

:end
pause
