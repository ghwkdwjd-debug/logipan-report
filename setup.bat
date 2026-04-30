@echo off
chcp 65001 > nul
title 로지판 - 처음 설치
cd /d "%~dp0"

echo ====================================================
echo   🚀 로지판 처음 설치
echo ====================================================
echo.

REM Python 설치 여부 확인
echo [1/4] Python 설치 확인 중...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [!] Python이 설치되어 있지 않습니다.
    echo.
    echo     아래 주소에서 Python을 먼저 설치해주세요:
    echo     https://www.python.org/downloads/
    echo.
    echo     ※ 설치할 때 반드시 "Add Python to PATH" 체크하세요!
    echo.
    pause
    exit /b 1
)
python --version
echo.

REM key.json 확인
echo [2/4] Firebase 키 파일 확인...
if not exist "key.json" (
    echo.
    echo [!] key.json 파일이 없습니다.
    echo     관리자에게 받은 key.json을 이 폴더에 넣어주세요:
    echo     %~dp0
    echo.
    pause
    exit /b 1
)
echo     ✅ key.json 확인 완료
echo.

REM 첫 실행해서 라이브러리 설치
echo [3/4] 라이브러리 설치 + 첫 실행...
echo.
python updater.py

REM 바탕화면에 박스 아이콘 바로가기 만들기
echo.
echo [4/4] 바탕화면에 박스 아이콘 바로가기 생성...

REM PowerShell로 .lnk 생성 (박스 아이콘 적용)
powershell -ExecutionPolicy Bypass -Command ^
  "$WshShell = New-Object -ComObject WScript.Shell;" ^
  "$Shortcut = $WshShell.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\로지판.lnk');" ^
  "$Shortcut.TargetPath = '%~dp0로지판.bat';" ^
  "$Shortcut.WorkingDirectory = '%~dp0';" ^
  "$Shortcut.IconLocation = '%~dp0로지판.ico,0';" ^
  "$Shortcut.WindowStyle = 7;" ^
  "$Shortcut.Description = '로지판 (Logi-Pan)';" ^
  "$Shortcut.Save()"

if exist "%USERPROFILE%\Desktop\로지판.lnk" (
    echo     ✅ 바탕화면에 박스 아이콘 바로가기 생성 완료
) else (
    echo     ⚠️ 바로가기 생성 실패. 직접 로지판.bat 사용하세요.
)

echo.
echo ====================================================
echo   ✅ 설치 완료!
echo ====================================================
echo.
echo   바탕화면의 [📦 로지판] 아이콘을 더블클릭해서 실행하세요.
echo.
pause
