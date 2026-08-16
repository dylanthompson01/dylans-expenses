@echo off
REM Local dev server for Dylan's Expenses.
REM Double-click this file, or run: serve.bat
REM Stop the server with Ctrl+C.

cd /d "%~dp0"

echo.
echo   Dylan's Expenses - local server
echo   http://localhost:8000
echo.
echo   Press Ctrl+C to stop.
echo.

start "" http://localhost:8000
python -m http.server 8000
