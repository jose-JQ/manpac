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

where nvidia-smi >nul 2>&1
if %errorlevel%==0 (
    echo GPU NVIDIA detectada. Intentando PyTorch con CUDA para DocTR...
    venv\Scripts\python.exe -m pip install --upgrade torch torchvision --index-url https://download.pytorch.org/whl/cu124
    if errorlevel 1 (
        echo No hay rueda CUDA para este Python; DocTR seguira en CPU. Ollama si puede usar la GPU.
    )
) else (
    echo No se detecto nvidia-smi: DocTR y el fallback de Ollama iran en CPU.
)

echo.
echo Listo. Ahora instala los programas externos (ver README.md):
echo   1. Ollama 0.7+  +  ollama pull qwen2.5:3b  +  ollama pull qwen2.5vl:3b
echo   2. Tesseract OCR
echo   3. Poppler
echo Luego ejecuta iniciar.bat
pause
