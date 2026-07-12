@echo off
set PY=%LOCALAPPDATA%\Programs\Python\Python313\python.exe
if not exist "%PY%" set PY=python
"%PY%" "%~dp0mukerrer_sil.py"
pause
