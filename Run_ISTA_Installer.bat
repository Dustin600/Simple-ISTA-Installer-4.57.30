@echo off
:: ============================================================
::  BMW ISTA 4.57.30 - Installer Launcher
::  Searches all drives for ISTA_Install.ps1 automatically
:: ============================================================

:: Check if already running as admin, if not relaunch as admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: First check same folder as this .bat
set "PS1_PATH=%~dp0ISTA_Install.ps1"
if exist "%PS1_PATH%" goto :RunScript

:: Not found next to bat - search all drives
echo ISTA_Install.ps1 not found next to launcher. Searching all drives...
set "PS1_PATH="

for %%D in (A B C D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if exist "%%D:\" (
        echo   Scanning %%D:\...
        for /f "delims=" %%F in ('dir /s /b "%%D:\ISTA_Install.ps1" 2^>nul') do (
            if not defined PS1_PATH set "PS1_PATH=%%F"
        )
    )
)

if not defined PS1_PATH (
    echo.
    echo ===============================================
    echo   ERROR: ISTA_Install.ps1 could not be found
    echo   on any drive. Please download it and place
    echo   it anywhere on your PC, then try again.
    echo ===============================================
    pause
    exit /b 1
)

echo   Found: %PS1_PATH%

:RunScript
echo.
echo   Launching: %PS1_PATH%
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1_PATH%"

pause
