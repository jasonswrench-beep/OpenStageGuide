@echo off
rem Starts OpenStage Guide on this computer and opens it in your browser. Close the server window to stop it.
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo OpenStage Guide needs Node.js 20 or later. Install it from https://nodejs.org and try again.
  pause
  exit /b 1
)
if not exist node_modules\acorn (
  echo First start: getting what OpenStage Guide needs...
  call npm ci --ignore-scripts
)
start "OpenStage Guide server" /min node scripts\project-server.mjs --port 3850
timeout /t 2 /nobreak >nul
start "" http://127.0.0.1:3850/
