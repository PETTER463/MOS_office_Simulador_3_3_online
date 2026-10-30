@echo off
setlocal
cd /d "%~dp0"
set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
if not defined PY for %%P in (py.exe python.exe) do (if not defined PY where %%P >nul 2>nul && set "PY=%%P")
if not defined PY (echo No se encontro Python 3.20+. & pause & exit /b 1)
if not exist ".venv\Scripts\python.exe" (
 "%PY%" -m venv .venv
 if errorlevel 1 (echo No se pudo crear entorno Python. & pause & exit /b 1)
)
".venv\Scripts\python.exe" -c "import docx,openpyxl,pptx" >nul 2>&1
if errorlevel 1 (
 echo Instalando dependencias en entorno aislado...
 ".venv\Scripts\python.exe" -m pip install -r requirements.txt
 if errorlevel 1 (echo Error instalando librerias. & pause & exit /b 1)
)
echo Iniciando MOS Office Simulator 3.2 en http://127.0.0.1:8773
".venv\Scripts\python.exe" server.py
if errorlevel 1 (echo Error al iniciar servidor. & pause)
