@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m PyInstaller --noconfirm --clean --windowed --name RejestrInterwencji app.py
if errorlevel 1 (
  echo.
  echo BLAD BUDOWANIA APLIKACJI.
  pause
  exit /b 1
)
echo.
echo GOTOWE. Aplikacja znajduje sie w folderze:
echo %~dp0dist\RejestrInterwencji\
echo.
echo Przenos caly folder RejestrInterwencji, nie sam plik EXE.
pause
