@echo off
setlocal
pushd "%~dp0"
set "batch_python=python"
if exist ".venv\Scripts\python.exe" set "batch_python=.venv\Scripts\python.exe"
if exist ".venv-build\Scripts\python.exe" set "batch_python=.venv-build\Scripts\python.exe"
if exist ".venv-build-customtkinter\Scripts\python.exe" set "batch_python=.venv-build-customtkinter\Scripts\python.exe"
"%batch_python%" -B "main.py"
if errorlevel 1 pause
popd
