@echo off
setlocal EnableDelayedExpansion
:: Check 365-Day MoonFlower Automation Schedule Status

echo ============================================================
echo MoonFlower 365-Day Schedule Status Check
echo ============================================================
echo.

:: Set project root
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

echo 📊 Checking current schedule status...
echo.

:: Run PowerShell status checker
powershell.exe -ExecutionPolicy Bypass -File "%PROJECT_ROOT%\Setup_365Day_Complete_Schedule.ps1" -Status

echo.
echo 💡 Tips:
echo   • All tasks should show "Ready" status
echo   • Next run times should be properly scheduled
echo   • If any task shows "Disabled", re-run the installer
echo.

pause
