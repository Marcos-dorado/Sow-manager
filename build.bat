@echo off
cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo.
    echo No se encontro el entorno virtual ^(carpeta venv\^).
    echo Antes de compilar, hay que crearlo una sola vez:
    echo.
    echo     python -m venv venv
    echo     venv\Scripts\activate
    echo     pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

call venv\Scripts\activate

pyinstaller --onedir --windowed ^
    --name "SOW Manager" ^
    --hidden-import "win32com.client" ^
    --hidden-import "pythoncom" ^
    --hidden-import "win32api" ^
    --hidden-import "pywintypes" ^
    --hidden-import "win32con" ^
    --hidden-import "win32timezone" ^
    --collect-all "openpyxl" ^
    --collect-all "docx" ^
    --collect-all "bs4" ^
    --collect-all "lxml" ^
    run.py

if errorlevel 1 (
    echo.
    echo La compilacion fallo. Revisa el error de arriba ^(por ejemplo, que
    echo pyinstaller este instalado: pip install -r requirements.txt^).
    pause
    exit /b 1
)

echo.
echo Build completado. Revisa dist\SOW Manager\
pause
