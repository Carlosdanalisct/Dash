@echo off
cd /d "%~dp0"
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8787" ^| findstr "LISTENING"') do taskkill /F /PID %%a >nul 2>nul
start "" "http://127.0.0.1:8787"
"C:\Users\Carlos\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "%~dp0app.py" --serve --host 127.0.0.1 --port 8787
pause
