@echo off
REM ============================================================
REM  build.bat  -  Compila Rename_Files.exe con PyInstaller
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
    --name "Rename_Files" ^
    Rename_Files.py

echo.
echo [3/3] Copiando .exe a la carpeta raiz...
if exist "dist\Rename_Files.exe" (
    copy /Y "dist\Rename_Files.exe" "Rename_Files.exe"
    echo.
    echo  OK  -  Rename_Files.exe generado correctamente.
) else (
    echo  ERROR - No se encontro el ejecutable en dist\
)

echo.
pause
