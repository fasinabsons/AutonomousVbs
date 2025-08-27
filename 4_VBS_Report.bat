@echo off
setlocal EnableDelayedExpansion
:: BAT 4: VBS Report Generation - SCHEDULED 4:00 PM
:: Phase 1 → Phase 4 ONLY
:: Scheduled at 4:00 PM after upload completion

echo ========================================
echo BAT 4: VBS Report Generation (4:00 PM)
echo ========================================
echo Current Time: %TIME%
echo Current Date: %DATE%
echo.

:: Enhanced time validation for 4:00 PM execution (4:00-4:05 PM window)
for /f %%i in ('powershell -Command "Get-Date -Format 'HHmm'"') do set "CURRENT_TIME=%%i"

echo 🕓 Current time: !CURRENT_TIME! (24-hour format)
echo 🎯 Expected execution window: 16:00-16:05 (4:00-4:05 PM)

if !CURRENT_TIME! LSS 1600 (
    echo ⏰ Too early - Report generation window is 4:00-4:05 PM
    echo [%TIME%] Report attempted outside time window >> %LOG_FILE%
    echo ℹ️ Waiting for upload process to complete
    pause
    exit /b 0
)

if !CURRENT_TIME! GTR 1605 (
    echo ⏰ Execution window passed - Report generation should be at 4:00-4:05 PM
    echo [%TIME%] Report attempted outside time window >> %LOG_FILE%
    echo ⚠️ Manual execution allowed
    echo.
    echo Continue anyway? (Press any key or Ctrl+C to cancel)
    pause
)

echo ✅ Time validation passed - executing VBS report generation
echo.

:: Set project root (automatically detect from BAT file location)
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

:: Create simple log file
set LOG_FILE=EHC_Logs\vbs_report_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting SIMPLE VBS Report Workflow >> %LOG_FILE%

echo.
echo 🚪 Ensuring VBS is closed before starting...
echo [%TIME%] Closing VBS before starting >> %LOG_FILE%
taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
taskkill /f /im "VBS.exe" /t >nul 2>&1  
taskkill /f /im "vbs*.exe" /t >nul 2>&1
taskkill /f /im "*absons*.exe" /t >nul 2>&1
timeout /t 2 /nobreak >nul
echo ✅ VBS closed before starting

echo.
echo 🔑 PHASE 1: VBS Login
echo ========================================
echo [%TIME%] Starting VBS Phase 1 >> %LOG_FILE%

python vbs\vbs_phase1_login.py
set PHASE1_EXIT=%errorlevel%

echo 📊 VBS Phase 1 exit code: %PHASE1_EXIT%
echo [%TIME%] VBS Phase 1 exit code: %PHASE1_EXIT% >> %LOG_FILE%

if !PHASE1_EXIT! neq 0 (
    echo ❌ VBS Phase 1 failed!
    echo [%TIME%] ERROR: VBS Phase 1 failed with code %PHASE1_EXIT% >> %LOG_FILE%
    exit /b 1
)
echo ✅ VBS Phase 1 completed successfully

echo.
echo 📊 PHASE 4: PDF Report Generation
echo ========================================
echo [%TIME%] Starting VBS Phase 4 >> %LOG_FILE%

python vbs\vbs_phase4_report_fixed.py
set PHASE4_EXIT=%errorlevel%

echo 📊 VBS Phase 4 exit code: %PHASE4_EXIT%
echo [%TIME%] VBS Phase 4 exit code: %PHASE4_EXIT% >> %LOG_FILE%

if !PHASE4_EXIT! neq 0 (
    echo ❌ VBS Phase 4 failed!
    echo [%TIME%] ERROR: VBS Phase 4 failed with code %PHASE4_EXIT% >> %LOG_FILE%
    exit /b 1
)
echo ✅ VBS Phase 4 completed successfully

echo.
echo 🚪 Ensuring VBS is closed after completion...
echo [%TIME%] Closing VBS after completion >> %LOG_FILE%
taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
taskkill /f /im "VBS.exe" /t >nul 2>&1  
taskkill /f /im "vbs*.exe" /t >nul 2>&1
taskkill /f /im "*absons*.exe" /t >nul 2>&1
timeout /t 2 /nobreak >nul
echo ✅ VBS closed after completion

echo.
echo 🎉 VBS Report Process Completed Successfully!
echo [%TIME%] VBS report workflow completed >> %LOG_FILE%
echo.
echo ✅ VBS Phase 1 (Login) completed
echo ✅ VBS Phase 4 (PDF Generation) completed
echo ✅ VBS Application closed properly
echo.
echo Report generation completed at %TIME% on %DATE%
echo ========================================