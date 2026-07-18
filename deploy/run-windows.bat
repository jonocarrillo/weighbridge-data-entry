@echo off
rem Weighbridge launcher with auto-restart on crash.
rem Logs to data\server.log
cd /d "%~dp0.."
:loop
echo [%date% %time%] starting server >> data\server.log
node server.js >> data\server.log 2>&1
echo [%date% %time%] server exited (code %errorlevel%), restarting in 3s >> data\server.log
timeout /t 3 /nobreak > nul
goto loop
