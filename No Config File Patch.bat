@echo off
title No Config File Patch - BMW ISTA 4.57.30
color 0B

echo.
echo  =============================================
echo   No Config File Patch - BMW ISTA 4.57.30
echo  =============================================
echo.
echo  This will apply the No Config File Patch to
echo  your ISTA installer, enabling support for
echo  both Rheingold and Modular ISTA layouts.
echo.
echo  Both files must be in the same folder as
echo  this launcher:
echo.
echo    - No Config File Patch.bat  (this file)
echo    - No Config File Patch.ps1  (the patch script)
echo.
pause

:: Check for admin rights - re-launch elevated if needed
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Requesting Administrator privileges...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Check the PS1 exists in the same folder
if not exist "%~dp0No Config File Patch.ps1" (
    echo.
    echo  [ERROR] No Config File Patch.ps1 not found.
    echo  Make sure both files are in the same folder.
    echo.
    pause
    exit /b 1
)

echo.
echo  Running patch as Administrator...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0No Config File Patch.ps1"

echo.
echo  Patch complete. Press any key to close.
pause >nul
