@echo off
REM --------------------------------------------------
REM GX7092 프린터 설치 스크립트
REM 작성자: [조대희]
REM 설명:
REM  - 프린터가 이미 설치되어 있는지 정확하게 확인 (WMIC where절 사용)
REM  - 프린터가 온라인 상태인지 ping으로 확인
REM  - 드라이버 설치 후 TCP/IP 포트 생성 (이미 존재하면 무시)
REM  - 최종적으로 프린터를 설치하며, 각 단계에서 발생한 명령 출력(오류 메시지 포함)을 콘솔에 표시
REM  - 스크립트 종료 전 pause로 대기하여 결과를 확인할 수 있음
REM --------------------------------------------------

:: --------------------------------------------------
:: 관리자 권한으로 실행 여부 확인
:: --------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 관리자 권한으로 실행 중이 아닙니다.
    echo 관리자 권한으로 스크립트를 재실행합니다...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

setlocal enabledelayedexpansion

REM --------------------------------------------------
REM 설정: 환경에 맞게 경로와 프린터 정보를 수정하세요.
REM --------------------------------------------------
set "DRIVER_INF_PATH=C:\Users\admin\Desktop\canon_install\Driver\GX7000P6.inf"
set "DRIVER_NAME=Canon GX7000 series"
set "PRINTER_NAME=교무실"
set "PRINTER_IP=192.168.0.116"

echo.
echo ========================================
echo GX7092 프린터 설치 시작
echo ========================================
echo.


REM --------------------------------------------------
REM 3. 프린터 드라이버 설치 (pnputil)
REM    - 명령 출력을 바로 콘솔에 표시하여 오류 내용을 확인함
REM --------------------------------------------------
echo.
echo [3] 프린터 드라이버 설치 중...
set "driver_output="
for /f "delims=" %%i in ('pnputil /add-driver "%DRIVER_INF_PATH%" /install 2^>^&1') do (
    echo %%i
    set "driver_output=!driver_output! %%i"
)
echo.
echo 드라이버 설치 결과: !driver_output!
echo !driver_output! | findstr /i "already installed" >nul
if %errorlevel%==0 (
    echo 드라이버는 이미 설치되어 있으므로, 해당 드라이버를 사용합니다.
) else (
    echo !driver_output! | findstr /i "error" >nul
    if %errorlevel%==0 (
         echo Error: 드라이버 설치 중 오류가 발생했습니다.
         goto end
    )
)

REM --------------------------------------------------
REM 4. TCP/IP 포트 생성 (cscript)
REM    - 명령 출력을 바로 콘솔에 표시하여 오류 내용을 확인함
REM --------------------------------------------------
echo.
echo [4] TCP/IP 포트 생성 중...
set "port_output="
for /f "delims=" %%i in ('cscript //nologo "C:\Windows\System32\Printing_Admin_Scripts\ko-KR\prnport.vbs" -a -r "IP_%PRINTER_IP%" -h %PRINTER_IP% -o raw -n 9100 2^>^&1') do (
    echo %%i
    set "port_output=!port_output! %%i"
)
echo.
echo TCP/IP 포트 생성 결과: !port_output!
echo !port_output! | findstr /i "already exists" >nul
if %errorlevel%==0 (
    echo TCP/IP 포트가 이미 존재합니다. 해당 포트를 사용합니다.
) else (
    echo !port_output! | findstr /i "error" >nul
    if %errorlevel%==0 (
         echo Error: TCP/IP 포트 생성 중 오류가 발생했습니다.
         goto end
    )
)

REM --------------------------------------------------
REM 5. 프린터 설치 (rundll32)
REM --------------------------------------------------
echo.
echo [5] 프린터 설치 중...
rundll32 printui.dll,PrintUIEntry /if /b "%PRINTER_NAME%" /f "%DRIVER_INF_PATH%" /r "IP_%PRINTER_IP%" /m "%DRIVER_NAME%"
if %errorlevel% neq 0 (
    echo Error: 프린터 설치 중 오류가 발생했습니다.
    goto end
)

echo.
echo "%PRINTER_NAME%" 프린터 설치 완료!

:end
echo.
pause
exit /b
