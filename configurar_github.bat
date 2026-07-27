@echo off
setlocal

set "APP_DIR=C:\GestionSilabos"
set "REPO_URL=https://github.com/RonaldTubay/Programa.git"

title Configurar GitHub GestionSilabos
echo ===============================================
echo   Configuracion de GitHub para GestionSilabos
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
  git init
)

git branch -M main
git remote remove origin >nul 2>nul
git remote add origin "%REPO_URL%"

echo.
echo Repositorio remoto configurado:
git remote -v
echo.
echo Ahora puedes ejecutar actualizar_sistema.bat.
pause
