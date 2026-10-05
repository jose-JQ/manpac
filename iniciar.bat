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

set "OLLAMA_HOST=http://127.0.0.1:11434"

echo Comprobando que Ollama no este expuesto en la red...
netstat -an | findstr /C:"0.0.0.0:11434" /C:"[::]:11434" >nul
if not errorlevel 1 (
    echo.
    echo AVISO: Ollama escucha en 0.0.0.0:11434. Cierra el puerto a la red.
    echo En la app de Ollama Windows desactiva "Expose Ollama to the network",
    echo define OLLAMA_HOST=127.0.0.1:11434 en variables de usuario y reinicia el servicio.
    echo.
)

echo Iniciando ManPAC (API + interfaz) en http://127.0.0.1:8000 ...
start "ManPAC" cmd /k "cd /d ""%~dp0"" && set OLLAMA_HOST=http://127.0.0.1:11434 && venv\Scripts\python.exe api_backend.py"

timeout /t 4 >nul
start http://127.0.0.1:8000

echo.
echo Interfaz y API: http://127.0.0.1:8000
echo Ollama cliente: %OLLAMA_HOST%
echo Cierra la ventana negra para detener el sistema.
pause
