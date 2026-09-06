@echo off
setlocal
cd /d "%~dp0"

echo.
echo === AI Product Assistant - Windows portable build ===
echo.

py -3.11 --version >nul 2>&1
if errorlevel 1 goto :NO_PYTHON

if not exist ".venv\Scripts\python.exe" (
    py -3.11 -m venv .venv
    if errorlevel 1 goto :FAILED
)

call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
if errorlevel 1 goto :FAILED

echo Installing OCR runtime. This can take several minutes on the first run...
python -m pip install paddlepaddle==3.2.0 -i https://www.paddlepaddle.org.cn/packages/stable/cpu/
if errorlevel 1 goto :FAILED

python -m pip install -r requirements.txt
if errorlevel 1 goto :FAILED

python -m pip install pyinstaller
if errorlevel 1 goto :FAILED

echo.
echo Packaging the portable folder...
REM Do not use --collect-all paddle: it scans optional JIT modules that can
REM crash the PyInstaller child process on Windows. PaddleOCR 3.x loads its
REM PaddleX OCR pipeline dynamically, so collect all PaddleX modules/resources
REM while collecting only Paddle binaries.
python -m PyInstaller --noconfirm --clean --windowed --onedir --name "AI_Product_Assistant" --distpath "dist_portable" --workpath "build_cache" --specpath "build_cache" --collect-data paddleocr --collect-all paddlex --collect-binaries paddle --collect-binaries cv2 --copy-metadata paddleocr --copy-metadata paddlex --copy-metadata paddlepaddle --copy-metadata beautifulsoup4 --copy-metadata einops --copy-metadata ftfy --copy-metadata imagesize --copy-metadata Jinja2 --copy-metadata latex2mathml --copy-metadata lxml --copy-metadata opencv-contrib-python --copy-metadata openpyxl --copy-metadata premailer --copy-metadata pyclipper --copy-metadata pypdfium2 --copy-metadata python-bidi --copy-metadata regex --copy-metadata safetensors --copy-metadata scikit-learn --copy-metadata scipy --copy-metadata sentencepiece --copy-metadata shapely --copy-metadata tiktoken --copy-metadata tokenizers ai_product_assistant.py
if errorlevel 1 goto :FAILED

REM Bundle pre-downloaded OCR models when they are available.  This keeps the
REM portable app usable on computers that cannot access model download sites.
if exist "_paddlex_cache\official_models\PP-OCRv5_mobile_det" (
    xcopy "_paddlex_cache" "dist_portable\AI_Product_Assistant\_paddlex_cache\" /E /I /Y >nul
    if errorlevel 1 goto :FAILED
)

echo.
echo Done.
echo Portable application folder:
echo %CD%\dist_portable\AI_Product_Assistant
echo.
echo Copy the entire AI_Product_Assistant folder to another Windows computer.
echo Do not copy only the EXE file.
pause
start "" "%CD%\dist_portable\AI_Product_Assistant"
exit /b 0

:NO_PYTHON
echo Python 3.11 (64-bit) was not found.
echo Install Python 3.11 from https://www.python.org/downloads/
echo During installation, enable "Add Python to PATH".
pause
exit /b 1

:FAILED
echo.
echo Build failed. Please take a screenshot of this window and send it to Codex.
pause
exit /b 1
