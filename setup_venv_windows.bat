::  run by     " .\setup_venv.bat "  in the terminal

@echo off
:: Check if python is installed
python --version >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python first.
    exit /b
)

:: Check if venv module is available (comes with Python >= 3.3)
python -m venv --help >nul 2>nul
if %errorlevel% neq 0 (
    echo Python venv module is not available. Ensure Python >= 3.3 is installed.
    exit /b
)

:: Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv .\venv
) else (
    echo virtual environment exist
)


:: Enable delayed variable expansion (this is for visual check-in – if dir is correct)
setlocal enabledelayedexpansion

:: Initialize the variable
set ACTIVATE_PATH=

:: Loop to get the full path of activate.bat and assign it to the variable
for %%I in (.\venv\Scripts\activate.bat) do (
    set ACTIVATE_PATH=%%~fI
)

:: Output the variable to check its value
echo Full path to activate.bat: !ACTIVATE_PATH!

:: Check if the venv activation file exists before trying to call it
if exist !ACTIVATE_PATH! (
    echo Activating virtual environment...
    call !ACTIVATE_PATH!
) else (
    echo The system cannot find the virtual environment activation script.
    exit /b
)

:: Install dependencies from requirements.txt
if exist requirements.txt (
    echo Installing dependencies from requirements.txt...
    pip install -r requirements.txt
) else (
    echo No requirements.txt found.
)

::   !!!
:: Deactivate the virtual environment after installation (optional)
deactivate
echo Virtual environment setup complete.
pause
