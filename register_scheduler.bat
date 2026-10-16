@echo off
chcp 65001 >nul
echo ========================================================
echo  [MarketSignal] 윈도우 작업 스케줄러 2종 자동 등록
echo ========================================================
echo.

:: 1. 관리자 권한 확인
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] 관리자 권한이 필요합니다.
    echo 이 파일을 마우스 우클릭하여 '관리자 권한으로 실행'을 눌러주세요.
    echo.
    pause
    exit /b 1
)

:: 2. 현재 PC의 pythonw.exe 경로 자동 탐색
set PYTHONW_PATH=
for /f "tokens=*" %%i in ('where pythonw 2^>nul') do (
    if not defined PYTHONW_PATH set "PYTHONW_PATH=%%i"
)

if not defined PYTHONW_PATH (
    echo [오류] pythonw.exe를 찾을 수 없습니다.
    echo Python이 정상 설치되어 있고 환경변수 PATH에 등록되어 있는지 확인해주세요.
    echo.
    pause
    exit /b 1
)

:: 3. 현재 스크립트 파일 경로 설정
set "SCRIPT_FILE=%~dp0market_signal_report.py"

if not exist "%SCRIPT_FILE%" (
    echo [오류] market_signal_report.py 파일을 찾을 수 없습니다.
    echo 파일이 배치 스크립트와 동일한 폴더에 있는지 확인해주세요.
    echo.
    pause
    exit /b 1
)

echo [인식된 파이썬 실행 경로]
echo %PYTHONW_PATH%
echo.
echo [인식된 스크립트 경로]
echo %SCRIPT_FILE%
echo.

:: 4. 2대 스케줄 등록 실행
echo [1/2] MarketSignal_Daemon (평일 08:50) 등록 중...
schtasks /create /tn "MarketSignal_Daemon" /tr "\"%PYTHONW_PATH%\" \"%SCRIPT_FILE%\" --daemon" /sc weekly /d MON,TUE,WED,THU,FRI /st 08:50 /f
if %errorlevel% equ 0 (
    echo  └ 성공: MarketSignal_Daemon 등록 완료
) else (
    echo  └ 실패: 등록 중 오류 발생
)
echo.

echo [2/2] MarketSignal_Closing (평일 15:40) 등록 중...
schtasks /create /tn "MarketSignal_Closing" /tr "\"%PYTHONW_PATH%\" \"%SCRIPT_FILE%\" --closing" /sc weekly /d MON,TUE,WED,THU,FRI /st 15:40 /f
if %errorlevel% equ 0 (
    echo  └ 성공: MarketSignal_Closing 등록 완료
) else (
    echo  └ 실패: 등록 중 오류 발생
)
echo.

echo ========================================================
echo  [등록 결과 확인]
echo ========================================================
schtasks /query /fo TABLE | findstr /i "MarketSignal"
echo.
echo 작업이 완료되었습니다. 아무 키나 누르면 창이 닫힙니다.
pause >nul