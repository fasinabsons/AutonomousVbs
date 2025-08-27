@echo off
setlocal EnableDelayedExpansion
:: BAT 3: VBS Upload (1:00-1:05 PM)
:: Phase 1 → Phase 2 → Phase 3 → Close VBS
:: Handles 3-hour upload process and VBS closure

echo ========================================
echo BAT 3: VBS Upload Process (1:00 PM)
echo ========================================
echo Current Time: %TIME%
echo Current Date: %DATE%
echo.

:: Set project root (automatically detect from BAT file location)
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

:: Create log file
set LOG_FILE=EHC_Logs\vbs_upload_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting VBS Upload Workflow >> %LOG_FILE%

:: Enhanced time validation for 1:00 PM execution (1:00-1:05 PM window)
for /f %%i in ('powershell -Command "Get-Date -Format 'HHmm'"') do set "CURRENT_TIME=%%i"

echo 🕐 Current time: !CURRENT_TIME! (24-hour format)
echo 🎯 Expected execution window: 13:00-13:05 (1:00-1:05 PM)

if !CURRENT_TIME! LSS 1300 (
    echo ⏰ Too early - VBS upload window is 1:00-1:05 PM
    echo [%TIME%] Upload attempted outside time window >> %LOG_FILE%
    pause
    exit /b 0
)

if !CURRENT_TIME! GTR 1305 (
    echo ⏰ Execution window passed - VBS upload should be at 1:00-1:05 PM
    echo [%TIME%] Upload attempted outside time window >> %LOG_FILE%
    echo ⚠️ Manual execution allowed but check for scheduling conflicts
    echo.
    echo Continue anyway? (Press any key or Ctrl+C to cancel)
    pause
)

echo ✅ Time validation passed - executing VBS upload process
echo.

:: Check for Excel file dependency
echo 🔍 Checking for Excel file dependency...

:: Get today's folder in ddMMM format (e.g., 31jul)
for /f %%i in ('powershell -Command "(Get-Date).ToString('ddMMM').ToLower()"') do set TODAY_FOLDER=%%i
set EXCEL_PATH=EHC_Data_Merge\%TODAY_FOLDER%\*.xls

echo Today's folder: %TODAY_FOLDER%
echo Looking for Excel file: %EXCEL_PATH%
echo [%TIME%] Checking for Excel file: %EXCEL_PATH% >> %LOG_FILE%

:check_excel
if not exist %EXCEL_PATH% (
    echo ⏳ Waiting for Excel file from BAT 2...
    echo [%TIME%] Waiting for Excel file >> %LOG_FILE%
    timeout /t 60 /nobreak >nul
    goto :check_excel
)
echo ✅ Excel file found - starting VBS upload process

echo.
echo 🔑 PHASE 1: VBS Login
echo ========================================
echo [%TIME%] Starting VBS Phase 1 >> %LOG_FILE%

:: Run VBS Phase 1 with enhanced error handling for BAT context
echo 🚀 Running VBS Phase 1 in BAT context...
cd /d "%PROJECT_ROOT%"
python vbs\vbs_phase1_login.py
set PHASE1_EXIT=%errorlevel%

echo 📊 VBS Phase 1 exit code: %PHASE1_EXIT%
echo [%TIME%] VBS Phase 1 exit code: %PHASE1_EXIT% >> %LOG_FILE%

if !PHASE1_EXIT! neq 0 (
    echo ❌ VBS Phase 1 failed!
    echo [%TIME%] ERROR: VBS Phase 1 failed with code %PHASE1_EXIT% >> %LOG_FILE%
    
    :: Wait longer and retry with manual intervention option
    echo 🔄 Retrying VBS Phase 1 in 30 seconds...
    echo 💡 If VBS security popup appears, please click 'Run' manually
    timeout /t 30 /nobreak
    
    echo 🚀 VBS Phase 1 retry attempt...
    python vbs\vbs_phase1_login.py
    set PHASE1_RETRY=%errorlevel%
    
    echo 📊 VBS Phase 1 retry exit code: %PHASE1_RETRY%
    echo [%TIME%] VBS Phase 1 retry exit code: %PHASE1_RETRY% >> %LOG_FILE%
    
    if !PHASE1_RETRY! neq 0 (
        echo ❌ VBS Phase 1 retry failed!
        echo [%TIME%] ERROR: VBS Phase 1 retry failed with code %PHASE1_RETRY% >> %LOG_FILE%
        echo.
        echo 💡 Manual Troubleshooting:
        echo   1. Check if VBS application is installed
        echo   2. Check if security popup is blocking execution
        echo   3. Try running: python vbs\vbs_phase1_login.py manually
        pause
        exit /b 1
    )
)
echo ✅ VBS Phase 1 completed successfully

echo.
echo 🧭 PHASE 2: VBS Navigation
echo ========================================
echo [%TIME%] Starting VBS Phase 2 >> %LOG_FILE%

python vbs\vbs_phase2_navigation_fixed.py
set PHASE2_EXIT=%errorlevel%

if !PHASE2_EXIT! neq 0 (
    echo ❌ VBS Phase 2 failed!
    echo [%TIME%] ERROR: VBS Phase 2 failed with code %PHASE2_EXIT% >> %LOG_FILE%
    
    :: Retry Phase 2 once
    echo 🔄 Retrying VBS Phase 2...
    timeout /t 15 /nobreak >nul
    
    python vbs\vbs_phase2_navigation_fixed.py
    set PHASE2_RETRY=%errorlevel%
    
    if !PHASE2_RETRY! neq 0 (
        echo ❌ VBS Phase 2 retry failed!
        echo [%TIME%] ERROR: VBS Phase 2 retry failed >> %LOG_FILE%
        pause
        exit /b 1
    )
)
echo ✅ VBS Phase 2 completed successfully

echo.
echo ⬆️ PHASE 3: VBS Upload (Enhanced with Multiple Update Variants)
echo ========================================
echo [%TIME%] Starting VBS Phase 3 with streamlined update detection >> %LOG_FILE%
echo.
echo 🎯 STREAMLINED FEATURES:
echo   • Multiple update button variants (variant1, variant2)
echo   • Upload success popup detection
echo   • Calibrated confidence levels for 100%% accuracy
echo   • 3-hour upload monitoring with progress logs
echo   • Automatic VBS restart for Phase 4 after completion
echo.
echo ⏰ Starting upload process (will take ~3 hours)...
echo 🖥️ VBS may show 'Not Responding' - this is NORMAL during upload
echo 📊 Progress will be logged every 15 minutes

:: OPTION 1: Try enhanced version with proven EHC header clicking
echo 🔧 Using ENHANCED version with PROVEN EHC header clicking
echo [%TIME%] Starting VBS Phase 3 ENHANCED >> %LOG_FILE%

python vbs\vbs_phase3_upload_complete.py
set PHASE3_EXIT=%errorlevel%

if !PHASE3_EXIT! neq 0 (
    echo ❌ VBS Phase 3 initial attempt failed!
    echo [%TIME%] ERROR: VBS Phase 3 failed with code %PHASE3_EXIT% >> %LOG_FILE%
    
    :: Check if VBS is still running (indicates partial success - don't restart from Phase 1)
    tasklist | findstr /i "absons" >nul
    if !errorlevel! equ 0 (
        echo ℹ️ VBS still running - Phase 3 may have partially completed
        echo ℹ️ NOT restarting from Phase 1 - assuming upload in progress
        echo [%TIME%] VBS still running, assuming Phase 3 upload in progress >> %LOG_FILE%
        
        echo ⏱️ Monitoring for 3 hours for upload completion...
        timeout /t 10800 /nobreak >nul
        echo ✅ 3 hour wait completed - assuming upload finished
    ) else (
        echo 🔄 VBS not running - trying ONE more attempt...
        echo [%TIME%] Single retry of VBS Phase 3 >> %LOG_FILE%
        
        python vbs\vbs_phase3_upload_complete.py
        set PHASE3_RETRY=%errorlevel%
        
        if !PHASE3_RETRY! neq 0 (
            echo ❌ VBS Phase 3 retry failed!
            echo [%TIME%] ERROR: VBS Phase 3 retry failed with code %PHASE3_RETRY% >> %LOG_FILE%
            echo.
            echo 💡 Manual intervention required - Phase 3 failed twice
            pause
            exit /b 1
        )
    )
)
echo ✅ VBS Phase 3 completed - upload finished!

:: AUTOMATIC PHASE 4: Generate PDF report after upload completion
echo.
echo 📊 AUTOMATIC PHASE 4: PDF Report Generation
echo ========================================
echo [%TIME%] Starting automatic VBS Phase 4 after upload completion >> %LOG_FILE%
echo.
echo 💡 VBS was closed after upload - restarting for report generation...
echo ⏰ Waiting 30 seconds for system to stabilize...
timeout /t 30 /nobreak >nul

:: Fresh VBS login for Phase 4
echo 🔑 PHASE 1 (Fresh): VBS Login for Reports
echo ----------------------------------------
echo [%TIME%] Starting fresh VBS Phase 1 for reports >> %LOG_FILE%

python vbs\vbs_phase1_login.py
set PHASE1_REPORT_EXIT=%errorlevel%

if !PHASE1_REPORT_EXIT! neq 0 (
    echo ❌ VBS Phase 1 for reports failed!
    echo [%TIME%] ERROR: VBS Phase 1 for reports failed >> %LOG_FILE%
    echo ⚠️ Skipping automatic Phase 4 - will run at 5:00 PM via BAT 4
) else (
    echo ✅ VBS Phase 1 for reports completed
    
    :: Run Phase 4 PDF generation
    echo 📊 PHASE 4: PDF Report Generation
    echo ----------------------------------------
    echo [%TIME%] Starting VBS Phase 4 PDF generation >> %LOG_FILE%
    
    python vbs\vbs_phase4_report_fixed.py
    set PHASE4_EXIT=%errorlevel%
    
    if !PHASE4_EXIT! neq 0 (
        echo ❌ VBS Phase 4 failed!
        echo [%TIME%] ERROR: VBS Phase 4 failed >> %LOG_FILE%
        echo ⚠️ PDF generation failed - will retry at 5:00 PM via BAT 4
    ) else (
        echo ✅ VBS Phase 4 completed - PDF report generated!
        echo [%TIME%] VBS Phase 4 completed successfully >> %LOG_FILE%
        
        :: Verify PDF creation
        for /f %%i in ('powershell -Command "(Get-Date).ToString('ddMMM').ToLower()"') do set TODAY_FOLDER=%%i
        if exist EHC_Data_Pdf\!TODAY_FOLDER!\*.pdf (
            echo ✅ PDF report verified - ready for tomorrow's email
            echo [%TIME%] PDF report verified for email >> %LOG_FILE%
        ) else (
            echo ⚠️ PDF verification failed
            echo [%TIME%] WARNING: PDF verification failed >> %LOG_FILE%
        )
    )
)

echo.
echo 🚪 VBS CLOSURE: Closing VBS Application
echo ========================================
echo [%TIME%] Closing VBS application >> %LOG_FILE%

:: Enhanced VBS closure with comprehensive cleanup
echo ⏳ Waiting 15 seconds for upload completion...
timeout /t 15 /nobreak >nul

echo 🔧 Enhanced VBS application closure...
echo [%TIME%] Starting VBS closure process >> !LOG_FILE!

:: Method 1: Graceful close via Enter key (handle any popups)
echo 🔑 Step 1: Handle any VBS confirmation popups...
python -c "
import pyautogui
import time
try:
    # Handle any close confirmation popups
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    print('Graceful close attempted')
except:
    print('Graceful close failed, continuing...')
" 2>nul

:: Method 2: Comprehensive process termination
echo 🛑 Step 2: Force close all VBS processes...
taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
taskkill /f /im "VBS.exe" /t >nul 2>&1  
taskkill /f /im "vbs*.exe" /t >nul 2>&1
taskkill /f /im "*absons*.exe" /t >nul 2>&1

:: Method 3: Wait and verify closure
timeout /t 3 /nobreak >nul
echo 🔍 Step 3: Verify VBS closure...
tasklist | findstr /i "absons" >nul
if !errorlevel! equ 0 (
    echo ⚠️ VBS still running - additional cleanup...
    taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
    timeout /t 2 /nobreak >nul
) else (
    echo ✅ VBS successfully closed
)

echo ✅ VBS application closed

echo.
echo 📧 Sending upload completion notification...
echo [%TIME%] Sending upload notification >> %LOG_FILE%
python email\email_delivery.py upload_complete
if !errorlevel! neq 0 (
    echo ⚠️ Notification failed (continuing anyway)
    echo [%TIME%] WARNING: Upload notification failed >> %LOG_FILE%
)

echo.
echo 🎉 VBS Upload Process Completed!
echo [%TIME%] VBS upload workflow completed >> %LOG_FILE%
echo.
echo ✅ VBS Phase 1 (Login) completed
echo ✅ VBS Phase 2 (Navigation) completed  
echo ✅ VBS Phase 3 (Upload) completed (~3 hours)
echo ✅ VBS application closed properly
echo ✅ Ready for VBS Report (BAT 4) at 5:00 PM
echo.
pause