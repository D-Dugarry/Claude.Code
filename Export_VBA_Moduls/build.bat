@echo off
REM ============================================================
REM  build.bat  -  Compila Export_VBA_Moduls.exe con PyInstaller
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
    --name "Export_VBA_Moduls" ^
    --hidden-import win32com.client ^
    --hidden-import win32com.server.util ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --collect-all pywin32 ^
    Export_VBA_Moduls.py

echo.
echo [3/3] Copiando .exe a la carpeta raiz...
if exist "dist\Export_VBA_Moduls.exe" (
    copy /Y "dist\Export_VBA_Moduls.exe" "Export_VBA_Moduls.exe"
    echo.
    echo  OK  -  Export_VBA_Moduls.exe generado correctamente.
) else (
    echo  ERROR - No se encontro el ejecutable en dist\
)

echo.
pause
