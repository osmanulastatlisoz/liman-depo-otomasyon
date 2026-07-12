@echo off
set PY=%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe
if not exist "%PY%" set PY=pythonw
start "" "%PY%" "%~dp0birlesik_stok_sorgu.py"
