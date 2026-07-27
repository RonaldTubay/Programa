@echo off
setlocal

set "APP_DIR=C:\GestionSilabos"
set "REPO_URL=https://github.com/RonaldTubay/Programa.git"

title Actualizar GestionSilabos
echo ===============================================
echo   Actualizacion del sistema GestionSilabos
echo ===============================================
echo.

cd /d "%APP_DIR%" || (
  echo No se encontro la carpeta %APP_DIR%.
  pause
  exit /b 1
)

where git >nul 2>nul || (
  echo Git no esta instalado o no esta en el PATH.
  echo Instala Git for Windows y vuelve a intentar.
  pause
  exit /b 1
)

if not exist ".git" (
  echo Esta carpeta no esta conectada a Git.
  echo Inicializando repositorio local...
  git init
  git remote add origin "%REPO_URL%"
  git branch -M main
) else (
  git remote set-url origin "%REPO_URL%"
)

echo.
echo Descargando cambios desde GitHub...
git fetch origin main
if errorlevel 1 (
  echo No se pudieron descargar cambios desde GitHub.
  pause
  exit /b 1
)

echo.
echo Aplicando cambios...
git pull --ff-only origin main
if errorlevel 1 (
  echo No se pudo aplicar la actualizacion automaticamente.
  echo Revisa si hay cambios locales o conflictos.
  pause
  exit /b 1
)

echo.
echo Actualizando dependencias...
if not exist ".\venv\Scripts\python.exe" (
  python -m venv venv
)
.\venv\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
.\venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo Validando dependencias...
.\venv\Scripts\python.exe -c "import flask, pandas, pyodbc, openpyxl, xlrd, lxml; print('Dependencias OK')"
if errorlevel 1 (
  echo La validacion de dependencias fallo.
  pause
  exit /b 1
)

echo.
echo Actualizacion finalizada correctamente.
echo Si el sistema ya estaba abierto, reinicia la tarea o vuelve a abrirlo.
pause
