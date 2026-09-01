@echo off
setlocal

REM One-line equivalent (run once Python + dependencies are already set up):
REM python -m PyInstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py

set "PYTHON_INSTALL_VERSION=3.12.7"
set "PYTHON_EXE="

call :find_python
if not "%PYTHON_EXE%"=="" goto :have_python

echo Python was not found on this PC. Attempting automatic installation...
call :install_python
call :find_python
if not "%PYTHON_EXE%"=="" goto :have_python

goto :nopython

:have_python
echo Using Python: %PYTHON_EXE%

echo [1/2] Installing required packages...
"%PYTHON_EXE%" -m pip install --upgrade pip
if errorlevel 1 goto :error
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
REM Looks for a real python.exe (never a .bat/.cmd shim, and never
REM the Microsoft Store "app execution alias" stub that Windows puts
REM on PATH at %LocalAppData%\Microsoft\WindowsApps\python.exe even
REM when Python isn't actually installed): first on PATH, then in
REM the usual per-user and system install locations. Sets PYTHON_EXE
REM to a full path (empty if nothing usable was found).
REM ---------------------------------------------------------------
:find_python
set "PYTHON_EXE="
for /f "delims=" %%P in ('where python.exe 2^>nul ^| findstr /I /V "WindowsApps"') do set "PYTHON_EXE=%%P"
if not "%PYTHON_EXE%"=="" goto :eof

if exist "%LocalAppData%\Programs\Python\Python313\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python313\python.exe"
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python312\python.exe"
if exist "%LocalAppData%\Programs\Python\Python311\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python311\python.exe"
if exist "%LocalAppData%\Programs\Python\Python310\python.exe" set "PYTHON_EXE=%LocalAppData%\Programs\Python\Python310\python.exe"
if exist "%ProgramFiles%\Python313\python.exe" set "PYTHON_EXE=%ProgramFiles%\Python313\python.exe"
if exist "%ProgramFiles%\Python312\python.exe" set "PYTHON_EXE=%ProgramFiles%\Python312\python.exe"
if exist "%ProgramFiles%\Python311\python.exe" set "PYTHON_EXE=%ProgramFiles%\Python311\python.exe"
if exist "%ProgramFiles%\Python310\python.exe" set "PYTHON_EXE=%ProgramFiles%\Python310\python.exe"
goto :eof

REM ---------------------------------------------------------------
REM Tries to install Python with no user interaction: winget first,
REM then falls back to downloading the official installer directly.
REM ---------------------------------------------------------------
:install_python
where winget >nul 2>nul
if errorlevel 1 goto :install_download

echo Trying winget install of Python %PYTHON_INSTALL_VERSION%...
winget install -e --id Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
call :refresh_path
where python.exe 2>nul | findstr /I /V "WindowsApps" >nul
if not errorlevel 1 goto :eof

:install_download
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
