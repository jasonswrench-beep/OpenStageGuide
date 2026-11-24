@echo off
rem Starts OpenStage Guide on this computer and opens it in your browser. Close the server window to stop it.
rem The browser is opened only when it is OpenStage Guide that answers on the port (another program may be using it; it is asked with curl.exe, or with
rem PowerShell where there is no curl.exe), and if the program
rem does not start, what it said is shown here.
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo OpenStage Guide needs Node.js 20 or later. Install it from https://nodejs.org and try again.
  pause
  exit /b 1
)
if not exist node_modules\acorn (
  echo First start: getting what OpenStage Guide needs...
  call npm ci --ignore-scripts --omit=dev
)
set "LOG=%TEMP%\openstage-guide-server.log"
if exist "%LOG%" del "%LOG%" >nul 2>nul
start "OpenStage Guide server" /min cmd /c ""node" "scripts\project-server.mjs" --port 3850 >"%LOG%" 2>&1"
set "HAVE_CURL=1"
where curl.exe >nul 2>nul
if errorlevel 1 set "HAVE_CURL="
set /a TRIES=0
:wait
if defined HAVE_CURL goto viacurl
powershell -NoProfile -NonInteractive -Command "try { if ((Invoke-WebRequest -UseBasicParsing -TimeoutSec 2 -Uri 'http://127.0.0.1:3850/health').Content -match 'OpenStage Guide') { exit 0 } else { exit 1 } } catch { exit 1 }" >nul 2>nul
goto probed
:viacurl
curl.exe -s --max-time 2 http://127.0.0.1:3850/health 2>nul | findstr /C:"OpenStage Guide" >nul
:probed
if not errorlevel 1 goto ready
set /a TRIES+=1
if %TRIES% GEQ 20 goto failed
timeout /t 1 /nobreak >nul
goto wait
:ready
start "" http://127.0.0.1:3850/
exit /b 0
:failed
echo.
echo OpenStage Guide did not start. Another program may be using port 3850, or Node.js could not run it.
if exist "%LOG%" (
  echo What the program said:
  type "%LOG%"
)
pause
exit /b 1
