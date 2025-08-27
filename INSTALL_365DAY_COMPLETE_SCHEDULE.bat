@echo off
setlocal EnableDelayedExpansion
:: Install 365-Day MoonFlower Automation Schedule
:: Handles locked/unlocked states, power management, and PC restarts

echo ============================================================
echo MoonFlower 365-Day Automation Schedule Installer
echo ============================================================
echo.
echo 🎯 This will install a complete 365-day automation schedule:
echo.
echo   📧 8:00 AM  - Morning email with PDF report to GM
echo   📥 9:30 AM  - Download CSV files (Morning)
echo   📊 12:30 PM - Download CSV + Excel merge (Afternoon)
echo   ⬆️ 1:00 PM  - VBS Upload (3-hour process)
echo   📋 4:00 PM  - VBS Report generation
echo   🚀 Startup  - Recovery check for missed schedules
echo.
echo 🔐 All tasks will run regardless of:
echo   ✅ PC locked or unlocked state
echo   ✅ Power plugged in or on battery
echo   ✅ User logged in or not
echo   ✅ Will wake PC from sleep if needed
echo.

:: Check for administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ ADMINISTRATOR PRIVILEGES REQUIRED
    echo.
    echo Please right-click this file and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo ✅ Administrator privileges confirmed
echo.

:: Set project root
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

echo 🔧 Project Root: %PROJECT_ROOT%
echo.

:: Create logs directory
if not exist "EHC_Logs" (
    mkdir EHC_Logs
    echo ✅ Created EHC_Logs directory
)

echo 🚀 Installing 365-Day Schedule...
echo.

:: Run PowerShell installer
powershell.exe -ExecutionPolicy Bypass -File "%PROJECT_ROOT%\Setup_365Day_Complete_Schedule.ps1" -Install

if %errorlevel% equ 0 (
    echo.
    echo 🎉 SUCCESS: 365-Day MoonFlower Automation Schedule Installed!
    echo.
    echo 📋 To check status: INSTALL_365DAY_COMPLETE_SCHEDULE.bat status
    echo 🗑️ To uninstall: INSTALL_365DAY_COMPLETE_SCHEDULE.bat uninstall
    echo.
    echo ⚠️ IMPORTANT: Keep this folder at the current location
    echo    Moving the folder will break the scheduled tasks!
    echo.
    echo 🔄 The system will now automatically:
    echo   • Send emails at 8:00 AM with yesterday's PDF
    echo   • Download data twice daily (9:30 AM, 12:30 PM)
    echo   • Upload data at 1:00 PM (3-hour process)
    echo   • Generate reports at 4:00 PM
    echo   • Recover from missed schedules on startup
    echo   • Run 365 days without intervention
    echo.
) else (
    echo ❌ INSTALLATION FAILED
    echo.
    echo Check the logs in EHC_Logs folder for details
    echo.
)

pause
