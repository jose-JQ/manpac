@echo off
setlocal
cd /d "%~dp0"

echo === ManPAC Extractor IA: instalacion ===
where python >nul 2>&1
if errorlevel 1 (
    echo No se encontro Python. Instala Python 3.11 o superior y vuelve a intentar.
    pause
    exit /b 1
)

if not exist venv (
    echo Creando entorno virtual...
    python -m venv venv
)

echo Instalando librerias Python...
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt

echo.
echo Listo. Ahora instala los programas externos (ver README.md):
echo   1. Ollama 0.7+  +  ollama pull qwen2.5:3b  +  ollama pull qwen2.5vl:3b
echo      Deja Ollama en 127.0.0.1:11434, sin exponer a la red.
echo   2. Tesseract OCR (idioma spa)
echo   3. Poppler
echo Luego ejecuta iniciar.bat
pause
