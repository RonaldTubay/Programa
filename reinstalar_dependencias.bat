@echo off
setlocal

set "APP_DIR=C:\GestionSilabos"

title Reinstalar dependencias GestionSilabos
echo ===============================================
echo   Reinstalacion limpia de dependencias
echo ===============================================
echo.

cd /d "%APP_DIR%" || (
  echo No se encontro la carpeta %APP_DIR%.
  pause
  exit /b 1
)

if exist ".\venv" (
  echo Eliminando entorno virtual anterior...
  rmdir /s /q ".\venv"
)

echo Creando entorno virtual...
python -m venv venv

echo Actualizando pip...
.\venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel

echo Instalando dependencias del proyecto...
.\venv\Scripts\python.exe -m pip install --no-cache-dir --force-reinstall -r requirements.txt

echo Validando dependencias...
.\venv\Scripts\python.exe -c "import flask, pandas, pyodbc, openpyxl, xlrd, lxml; print('Dependencias OK')"

echo.
echo Proceso terminado.
pause
