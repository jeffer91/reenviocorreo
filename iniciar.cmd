@echo off
setlocal
set "venv=%LOCALAPPDATA%\ReenvioCorreo\python"
if not exist "%venv%\Scripts\python.exe" (
    where py >nul 2>nul
    if errorlevel 1 (
        where python >nul 2>nul
        if errorlevel 1 (
            echo Instala Python 3 para Windows y vuelve a iniciar.
            exit /b 1
        )
        python -m venv "%venv%"
    ) else (
        py -3 -m venv "%venv%"
    )
)
if not exist "%venv%\Scripts\python.exe" exit /b 1
"%venv%\Scripts\python.exe" -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    "%venv%\Scripts\python.exe" -m pip install pywin32
    if errorlevel 1 exit /b 1
)
"%venv%\Scripts\python.exe" -c "import tkinter" >nul 2>nul
if errorlevel 1 (
    echo Instala el componente Tcl/Tk de Python para abrir la interfaz.
    pause
    exit /b 1
)
start "" "%venv%\Scripts\pythonw.exe" "%~dp0reenviar.py"
exit /b %errorlevel%
