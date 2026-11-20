@echo off
REM Compila IAPAY_Atencion_Comercio.exe con un solo doble clic.
REM Requiere Python 3.10+ SOLO en la PC donde se compila (la PC final no necesita nada).
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo No se encontro Python en esta PC. Use la opcion GitHub Actions del README.
  pause & exit /b 1
)
py -m venv .venv_build || goto :err
call .venv_build\Scripts\activate.bat
python -m pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt pyinstaller || goto :err
pyinstaller --noconfirm --onefile --windowed --name IAPAY_Atencion_Comercio app.py || goto :err
echo.
echo Listo: dist\IAPAY_Atencion_Comercio.exe
pause
exit /b 0
:err
echo Hubo un error en la compilacion.
pause
exit /b 1
