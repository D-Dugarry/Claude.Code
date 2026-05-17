@echo off
echo.
echo ============================================
echo   GENERANDO ComparadorVBA.exe ...
echo ============================================
echo.
cd /d "F:\__Dugarry UA\Dugarry Proyectos\___Claude.Code\Comparador_VBA"

"C:\Users\FX706\AppData\Local\Programs\Python\Python314\python.exe" -m pip install pyinstaller --quiet

"C:\Users\FX706\AppData\Local\Programs\Python\Python314\python.exe" -m PyInstaller --onefile --windowed --name ComparadorVBA `
  --hidden-import win32com.client `
  --hidden-import win32api `
  --hidden-import pythoncom `
  --hidden-import pywintypes `
  --add-data "compare.py;." `
  --add-data "apply.py;." `
  gui.py

echo.
if exist dist\ComparadorVBA.exe (
  echo EXE generado en: F:\__Dugarry UA\Dugarry Proyectos\___Claude.Code\Comparador_VBA\dist\ComparadorVBA.exe
) else (
  echo ERROR: No se genero el exe. Revisa los mensajes anteriores.
)
echo.
pause