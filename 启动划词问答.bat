@echo off
setlocal
cd /d "%~dp0"
set "PYW=C:\Users\wj527129330\AppData\Local\Programs\Python\Python312\pythonw.exe"
if not exist "%PYW%" set "PYW=pythonw"
start "" "%PYW%" "%~dp0select_ai.py"
