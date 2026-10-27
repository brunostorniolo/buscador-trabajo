@echo off
REM Corre el buscador y guarda lo que muestra en historial.log
cd /d "%~dp0"
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" buscador.py --exportar >> historial.log 2>&1
