@echo off
setlocal

REM One-line equivalent (run after dependencies are installed):
REM pyinstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py

echo [1/2] Installing required packages...
pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/2] Building single-file executable with PyInstaller...
pyinstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py
if errorlevel 1 goto :error

echo.
echo Build complete: dist\PPTX2Markdown.exe
pause
goto :eof

:error
echo.
echo Build failed. See the output above for details.
pause
