@echo off
REM ============================================================
REM  build.bat  -  Compila Mailing_Personalizado.exe con PyInstaller
REM ============================================================

cd /d "%~dp0"

echo.
echo [1/3] Instalando dependencias...
python -m pip install -r requirements.txt --quiet

echo.
echo [2/3] Compilando .exe ...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "Mailing_Personalizado" ^
    --add-data "Mail.png;." ^
    --add-data "MailSent.gif;." ^
    Mailing_Personalizado.py

echo.
echo [3/3] Copiando .exe a la carpeta raiz...
if exist "dist\Mailing_Personalizado.exe" (
    copy /Y "dist\Mailing_Personalizado.exe" "Mailing_Personalizado.exe"
    echo.
    echo  OK  -  Mailing_Personalizado.exe generado correctamente.
) else (
    echo  ERROR - No se encontro el ejecutable en dist\
)

echo.
pause
