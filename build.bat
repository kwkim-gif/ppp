@echo off
setlocal

REM 동일 작업을 한 줄로 실행하려면 (패키지 설치는 별도로 완료된 상태여야 함):
REM pyinstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py

echo [1/2] 필요한 패키지를 설치합니다...
pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/2] PyInstaller로 단일 실행 파일(.exe)을 빌드합니다...
pyinstaller --onefile --windowed --noconfirm --name PPTX2Markdown --collect-all tkinterdnd2 app.py
if errorlevel 1 goto :error

echo.
echo 빌드 완료: dist\PPTX2Markdown.exe
pause
goto :eof

:error
echo.
echo 빌드 중 오류가 발생했습니다.
pause
