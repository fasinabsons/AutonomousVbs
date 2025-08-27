@echo off
setlocal EnableDelayedExpansion
:: Uninstall 365-Day MoonFlower Automation Schedule

echo ============================================================
echo MoonFlower 365-Day Schedule Uninstaller
echo ============================================================
echo.

echo ⚠️ WARNING: This will remove ALL scheduled automation tasks:
echo.
echo   📥 Morning downloads (9:30 AM)
echo   📊 Afternoon downloads + Excel merge (12:30 PM)
echo   ⬆️ VBS uploads (1:00 PM)
echo   🔄 PC restarts (2:00 PM)
echo   📋 Report generation (4:00 PM)
echo   📧 Email delivery (8:00 PM)
echo.

set /p CONFIRM="Are you sure you want to uninstall? (y/N): "
if /i not "%CONFIRM%"=="y" (
    echo Uninstall cancelled.
    pause
    exit /b 0
)

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

echo 🗑️ Uninstalling 365-Day Schedule...
echo.

:: Run PowerShell uninstaller
powershell.exe -ExecutionPolicy Bypass -File "%PROJECT_ROOT%\Setup_365Day_Complete_Schedule.ps1" -Uninstall

if %errorlevel% equ 0 (
    echo.
    echo 🎉 SUCCESS: All MoonFlower scheduled tasks removed!
    echo.
    echo ℹ️ The automation scripts are still available for manual execution
    echo 🔧 To reinstall: Run INSTALL_365DAY_COMPLETE_SCHEDULE.bat
    echo.
) else (
    echo ❌ UNINSTALL FAILED
    echo.
    echo Some tasks may not have been removed. Check Task Scheduler manually.
    echo.
)

pause
