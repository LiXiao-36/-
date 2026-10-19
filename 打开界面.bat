@echo off
setlocal
cd /d "%~dp0"
set "PYW=C:\Users\yungz\.dsh\dsh-runtimes\dsh-primary-runtime\dependencies\python\pythonw.exe"
if not exist "%PYW%" set "PYW=pythonw"
start "" "%PYW%" "%~dp0gui.py"
