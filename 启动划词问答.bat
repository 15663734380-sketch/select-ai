@echo off
setlocal
cd /d "%~dp0"
set "PYW="
where pythonw >nul 2>nul && set "PYW=pythonw"
if not defined PYW (
  for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%D\pythonw.exe" set "PYW=%%D\pythonw.exe"
  )
)
if not defined PYW (
  echo Python not found. Install Python 3 and try again.
  pause
  exit /b 1
)
start "" "%PYW%" "%~dp0select_ai.py"
