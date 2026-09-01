@echo off
setlocal

REM One-line equivalent (run once Python + dependencies are already set up):
REM python -m PyInstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py

set "PYTHON_INSTALL_VERSION=3.12.7"

call :find_python
if "%PYTHON_EXE%"=="" (
    echo Python was not found on this PC. Attempting automatic installation...
    call :install_python
    call :find_python
)

if "%PYTHON_EXE%"=="" goto :nopython

echo Using Python: %PYTHON_EXE%

echo [1/2] Installing required packages...
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/2] Building single-file executable with PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py
if errorlevel 1 goto :error

echo.
echo Build complete: dist\PPTX2Markdown.exe
pause
goto :eof

:nopython
echo.
echo Automatic Python installation failed.
echo Please install Python manually from https://www.python.org/downloads/
echo (check "Add python.exe to PATH" during setup), then run this script again.
pause
goto :eof

:error
echo.
echo Build failed. See the output above for details.
pause
goto :eof

REM ---------------------------------------------------------------
REM Looks for a usable python.exe: first on PATH, then in the usual
REM per-user and system install locations. Sets PYTHON_EXE (empty if
REM nothing was found).
REM ---------------------------------------------------------------
:find_python
set "PYTHON_EXE="
where python >nul 2>nul
if not errorlevel 1 (
    set "PYTHON_EXE=python"
    goto :eof
)
for %%V in (Python313 Python312 Python311 Python310) do (
    if exist "%LocalAppData%\Programs\Python\%%V\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\%%V\python.exe"
)
for %%V in (Python313 Python312 Python311 Python310) do (
    if exist "%ProgramFiles%\%%V\python.exe" set "PYTHON_EXE=%ProgramFiles%\%%V\python.exe"
)
goto :eof

REM ---------------------------------------------------------------
REM Tries to install Python with no user interaction: winget first,
REM then falls back to downloading the official installer directly.
REM ---------------------------------------------------------------
:install_python
where winget >nul 2>nul
if not errorlevel 1 (
    echo Trying winget install of Python %PYTHON_INSTALL_VERSION%...
    winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    call :refresh_path
)

where python >nul 2>nul
if not errorlevel 1 goto :eof

echo Downloading the official Python installer...
set "PY_INSTALLER=%TEMP%\python-installer.exe"
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/%PYTHON_INSTALL_VERSION%/python-%PYTHON_INSTALL_VERSION%-amd64.exe' -OutFile '%PY_INSTALLER%' } catch { exit 1 }"
if errorlevel 1 goto :eof
if not exist "%PY_INSTALLER%" goto :eof

echo Installing Python %PYTHON_INSTALL_VERSION% silently, please wait...
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0
call :refresh_path
goto :eof

REM ---------------------------------------------------------------
REM Pulls the current user + system PATH from the registry into this
REM script's PATH, so a PATH change made by an installer that just
REM ran becomes visible without opening a new command prompt.
REM ---------------------------------------------------------------
:refresh_path
set "USERPATH="
set "SYSPATH="
for /f "skip=2 tokens=2,*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USERPATH=%%B"
for /f "skip=2 tokens=2,*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSPATH=%%B"
set "PATH=%SYSPATH%;%USERPATH%;%PATH%"
goto :eof
