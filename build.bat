@echo off
cd /d "%~dp0"
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
echo.
echo Build completado. Revisa dist\SOW Manager\
pause