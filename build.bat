@echo off
setlocal

REM One-line equivalent (run after dependencies are installed):
REM python -m PyInstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py

where python >nul 2>nul
if errorlevel 1 goto :nopython

echo [1/2] Installing required packages...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/2] Building single-file executable with PyInstaller...
python -m PyInstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py
if errorlevel 1 goto :error

echo.
echo Build complete: dist\PPTX2Markdown.exe
pause
goto :eof

:nopython
echo.
echo Python was not found on PATH.
echo Install Python from https://www.python.org/downloads/ and make sure
echo "Add python.exe to PATH" is checked during setup, then run this script again.
pause
goto :eof

:error
echo.
echo Build failed. See the output above for details.
pause
