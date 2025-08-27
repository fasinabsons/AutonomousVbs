@echo off
:: BAT 6: Master Scheduler (Windows Task Scheduler Setup)
:: Sets up all 4 BAT files for automated 365-day operation
:: Requires Administrator privileges

echo ========================================
echo BAT 6: Master Scheduler Setup
echo ========================================
echo.

:: Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ This script requires Administrator privileges!
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

:: Set project root (automatically detect from BAT file location)
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%

echo ✅ Administrator privileges confirmed
echo 📁 Project Root: %PROJECT_ROOT%
echo.

echo 🕘 MOONFLOWER 365-DAY AUTOMATION SCHEDULE
echo ================================================================
echo.
echo Creating Windows Task Scheduler entries...
echo.

:: BAT 1: Morning Email (9:00-9:30 AM daily)
echo 📧 Setting up BAT 1: Morning Email (9:00 AM daily)
schtasks /create /tn "MoonFlower_Email_Morning" /tr "\"%PROJECT_ROOT%\1_Email_Morning.bat\"" /sc daily /st 09:00 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ Morning Email task scheduled
) else (
    echo ❌ Failed to schedule Morning Email task
)

:: BAT 2: Download Files - Morning (9:30 AM daily)
echo 📥 Setting up BAT 2a: File Downloads Morning (9:30 AM daily)
schtasks /create /tn "MoonFlower_Downloads_Morning" /tr "\"%PROJECT_ROOT%\2_Download_Files.bat\"" /sc daily /st 09:30 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ Morning Downloads task scheduled
) else (
    echo ❌ Failed to schedule Morning Downloads task
)

:: BAT 2: Download Files - Afternoon (12:30 PM daily)
echo 📥 Setting up BAT 2b: File Downloads Afternoon (12:30 PM daily)
schtasks /create /tn "MoonFlower_Downloads_Afternoon" /tr "\"%PROJECT_ROOT%\2_Download_Files.bat\"" /sc daily /st 12:30 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ Afternoon Downloads task scheduled
) else (
    echo ❌ Failed to schedule Afternoon Downloads task
)

:: BAT 2: Excel Merge (12:35 PM daily)
echo 📊 Setting up BAT 2c: Excel Merge (12:35 PM daily)
schtasks /create /tn "MoonFlower_Excel_Merge" /tr "\"%PROJECT_ROOT%\2_Download_Files.bat\"" /sc daily /st 12:35 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ Excel Merge task scheduled
) else (
    echo ❌ Failed to schedule Excel Merge task
)

:: BAT 3: VBS Upload (1:00 PM daily)
echo ⬆️ Setting up BAT 3: VBS Upload (1:00 PM daily)
schtasks /create /tn "MoonFlower_VBS_Upload" /tr "\"%PROJECT_ROOT%\3_VBS_Upload.bat\"" /sc daily /st 13:00 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ VBS Upload task scheduled
) else (
    echo ❌ Failed to schedule VBS Upload task
)

:: BAT 4: VBS Report (5:00 PM daily)
echo 📊 Setting up BAT 4: VBS Report (5:00 PM daily)
schtasks /create /tn "MoonFlower_VBS_Report" /tr "\"%PROJECT_ROOT%\4_VBS_Report.bat\"" /sc daily /st 17:00 /ru Lenovo /rl HIGHEST /f
if %errorlevel% equ 0 (
    echo ✅ VBS Report task scheduled
) else (
    echo ❌ Failed to schedule VBS Report task
)

:: Cleanup old tasks
echo.
echo 🧹 Cleaning up old automation tasks...
schtasks /delete /tn "MoonFlower_Complete_Daily" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_VBS_Only" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_Health_Check" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_Evening_Lock" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_DataCollection" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_VBS_Automation" /f >nul 2>&1
schtasks /delete /tn "MoonFlower_MorningEmail" /f >nul 2>&1
echo ✅ Old tasks cleaned up

echo.
echo ========================================
echo 365-Day Master Schedule Created!
echo ========================================
echo.
echo 📅 DAILY AUTOMATION SCHEDULE:
echo ----------------------------------------------------------------
echo   09:00 AM - BAT 1: Morning Email
echo              ↳ Send yesterday's PDF to GM
echo              ↳ Only if PDF exists
echo.
echo   09:30 AM - BAT 2: File Downloads (Morning)
echo              ↳ CSV downloads (4 files expected)
echo              ↳ Creates daily folders
echo.
echo   12:30 PM - BAT 2: File Downloads (Afternoon)
echo              ↳ Additional CSV downloads if needed
echo              ↳ Ensures all slots are covered
echo.
echo   12:35 PM - BAT 2: Excel Merge
echo              ↳ Merge CSV files into Excel
echo              ↳ Validate 6000+ rows
echo.
echo   01:00 PM - BAT 3: VBS Upload
echo              ↳ Phase 1 (Login)
echo              ↳ Phase 2 (Navigation)
echo              ↳ Phase 3 (Upload - 3 hours)
echo              ↳ Close VBS application
echo.
echo   05:00 PM - BAT 4: VBS Report
echo              ↳ Phase 1 (Fresh login)
echo              ↳ Phase 4 (PDF generation)
echo              ↳ Close VBS application
echo.
echo ✅ AUTOMATION FEATURES:
echo   • 365-day continuous operation
echo   • Automatic retry mechanisms
echo   • Multiple update button variants
echo   • Comprehensive error handling
echo   • Email notifications for key events
echo   • Runs with HIGHEST privileges
echo   • Works when user is logged on or not
echo.
echo 🔧 MANUAL CONTROLS:
echo   • Run individual BAT files for testing
echo   • BAT 5 for complete workflow testing
echo   • All tasks use user: Lenovo
echo   • All logs stored in EHC_Logs\[date]\ folders
echo.
echo 📋 TASK MANAGEMENT:
echo   View tasks: schtasks /query | findstr "MoonFlower"
echo   Run manually: schtasks /run /tn "MoonFlower_[TaskName]"
echo   Delete task: schtasks /delete /tn "MoonFlower_[TaskName]" /f
echo.
echo 🎯 SYSTEM REQUIREMENTS:
echo   • Keep laptop plugged in and running
echo   • VBS application accessible
echo   • Network connectivity for downloads
echo   • Outlook configured for GM emails
echo.

:: Test schedule creation
echo.
echo 🧪 Testing schedule setup...
schtasks /query /tn "MoonFlower_Email_Morning" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Schedule verification passed
) else (
    echo ❌ Schedule verification failed
)

echo.
echo 🎉 Ready for 365-day continuous operation!
echo Keep the laptop running and the automation will handle everything.
echo.
pause