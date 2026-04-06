@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Ambiente virtual nao encontrado.
  exit /b 1
)

".venv\Scripts\python.exe" -m PyInstaller ^
  --noconsole ^
  --onefile ^
  --name CloudBillingAgent ^
  --add-data "app\\mappings\\service_mapping.csv;app\\mappings" ^
  web_launcher.pyw
