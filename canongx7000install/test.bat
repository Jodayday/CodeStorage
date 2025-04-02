@echo off
REM --------------------------------------------------
REM GX7092 ������ ��ġ ��ũ��Ʈ
REM �ۼ���: [������]
REM ����:
REM  - �����Ͱ� �̹� ��ġ�Ǿ� �ִ��� ��Ȯ�ϰ� Ȯ�� (WMIC where�� ���)
REM  - �����Ͱ� �¶��� �������� ping���� Ȯ��
REM  - ����̹� ��ġ �� TCP/IP ��Ʈ ���� (�̹� �����ϸ� ����)
REM  - ���������� �����͸� ��ġ�ϸ�, �� �ܰ迡�� �߻��� ��� ���(���� �޽��� ����)�� �ֿܼ� ǥ��
REM  - ��ũ��Ʈ ���� �� pause�� ����Ͽ� ����� Ȯ���� �� ����
REM --------------------------------------------------

:: --------------------------------------------------
:: ������ �������� ���� ���� Ȯ��
:: --------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ������ �������� ���� ���� �ƴմϴ�.
    echo ������ �������� ��ũ��Ʈ�� ������մϴ�...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

setlocal enabledelayedexpansion

REM --------------------------------------------------
REM ����: ȯ�濡 �°� ��ο� ������ ������ �����ϼ���.
REM --------------------------------------------------
set "DRIVER_INF_PATH=C:\Users\admin\Desktop\canon_install\Driver\GX7000P6.inf"
set "DRIVER_NAME=Canon GX7000 series"
set "PRINTER_NAME=������"
set "PRINTER_IP=192.168.0.116"

echo.
echo ========================================
echo GX7092 ������ ��ġ ����
echo ========================================
echo.


REM --------------------------------------------------
REM 3. ������ ����̹� ��ġ (pnputil)
REM    - ��� ����� �ٷ� �ֿܼ� ǥ���Ͽ� ���� ������ Ȯ����
REM --------------------------------------------------
echo.
echo [3] ������ ����̹� ��ġ ��...
set "driver_output="
for /f "delims=" %%i in ('pnputil /add-driver "%DRIVER_INF_PATH%" /install 2^>^&1') do (
    echo %%i
    set "driver_output=!driver_output! %%i"
)
echo.
echo ����̹� ��ġ ���: !driver_output!
echo !driver_output! | findstr /i "already installed" >nul
if %errorlevel%==0 (
    echo ����̹��� �̹� ��ġ�Ǿ� �����Ƿ�, �ش� ����̹��� ����մϴ�.
) else (
    echo !driver_output! | findstr /i "error" >nul
    if %errorlevel%==0 (
         echo Error: ����̹� ��ġ �� ������ �߻��߽��ϴ�.
         goto end
    )
)

REM --------------------------------------------------
REM 4. TCP/IP ��Ʈ ���� (cscript)
REM    - ��� ����� �ٷ� �ֿܼ� ǥ���Ͽ� ���� ������ Ȯ����
REM --------------------------------------------------
echo.
echo [4] TCP/IP ��Ʈ ���� ��...
set "port_output="
for /f "delims=" %%i in ('cscript //nologo "C:\Windows\System32\Printing_Admin_Scripts\ko-KR\prnport.vbs" -a -r "IP_%PRINTER_IP%" -h %PRINTER_IP% -o raw -n 9100 2^>^&1') do (
    echo %%i
    set "port_output=!port_output! %%i"
)
echo.
echo TCP/IP ��Ʈ ���� ���: !port_output!
echo !port_output! | findstr /i "already exists" >nul
if %errorlevel%==0 (
    echo TCP/IP ��Ʈ�� �̹� �����մϴ�. �ش� ��Ʈ�� ����մϴ�.
) else (
    echo !port_output! | findstr /i "error" >nul
    if %errorlevel%==0 (
         echo Error: TCP/IP ��Ʈ ���� �� ������ �߻��߽��ϴ�.
         goto end
    )
)

REM --------------------------------------------------
REM 5. ������ ��ġ (rundll32)
REM --------------------------------------------------
echo.
echo [5] ������ ��ġ ��...
rundll32 printui.dll,PrintUIEntry /if /b "%PRINTER_NAME%" /f "%DRIVER_INF_PATH%" /r "IP_%PRINTER_IP%" /m "%DRIVER_NAME%"
if %errorlevel% neq 0 (
    echo Error: ������ ��ġ �� ������ �߻��߽��ϴ�.
    goto end
)

echo.
echo "%PRINTER_NAME%" ������ ��ġ �Ϸ�!

:end
echo.
pause
exit /b
