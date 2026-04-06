@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Ambiente virtual nao encontrado. Crie com:
  echo python -m venv .venv
  echo .venv\Scripts\pip install -r requirements.txt
  exit /b 1
)

".venv\Scripts\python.exe" -m app.web
