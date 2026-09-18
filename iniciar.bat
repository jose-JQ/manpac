@echo off
setlocal
cd /d "%~dp0"

if not exist venv\Scripts\python.exe (
    echo Primero ejecuta instalar.bat
    pause
    exit /b 1
)

if defined POPPLER_PATH (
    echo POPPLER_PATH=%POPPLER_PATH%
) else (
    set "POPPLER_PATH=C:\poppler-26.07.0\Library\bin"
)
if defined TESSERACT_CMD (
    echo TESSERACT_CMD=%TESSERACT_CMD%
) else (
    set "TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe"
)

echo Iniciando API en http://127.0.0.1:8000 ...
start "ManPAC API" cmd /k "cd /d ""%~dp0"" && venv\Scripts\python.exe api_backend.py"

timeout /t 4 >nul

echo Iniciando interfaz Streamlit en http://localhost:8501 ...
start "ManPAC UI" cmd /k "cd /d ""%~dp0"" && venv\Scripts\streamlit.exe run app.py"

echo.
echo Navegador: http://localhost:8501
echo API:       http://127.0.0.1:8000/api/health
echo Cierra las dos ventanas negras para detener el sistema.
pause
