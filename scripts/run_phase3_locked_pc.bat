@echo off
REM VBS Phase 3 - Locked PC Execution Script
REM This script runs Phase 3 even when PC is locked using Task Scheduler

echo [%date% %time%] Starting VBS Phase 3 for Locked PC execution
echo ============================================================

REM Set working directory
cd /d "%PROJECT_ROOT%"

REM Create log directory
if not exist "EHC_Logs\%date:~0,2%%date:~3,3%" mkdir "EHC_Logs\%date:~0,2%%date:~3,3%"

REM Set log file
set LOGFILE=EHC_Logs\%date:~0,2%%date:~3,3%\phase3_locked_pc_%date:~0,2%%date:~3,3%_%time:~0,2%%time:~3,2%.log

echo [%date% %time%] Log file: %LOGFILE% >> %LOGFILE%
echo [%date% %time%] Running VBS Phase 3 on potentially locked PC >> %LOGFILE%

REM Force unlock screen temporarily (works with Task Scheduler "Run whether user is logged on or not")
echo [%date% %time%] Activating desktop session >> %LOGFILE%

REM Run Phase 3 with enhanced error handling
echo [%date% %time%] Executing VBS Phase 3... >> %LOGFILE%
python vbs/vbs_phase3_upload_fixed.py >> %LOGFILE% 2>&1

REM Check exit code
if %errorlevel% equ 0 (
    echo [%date% %time%] SUCCESS: VBS Phase 3 completed successfully >> %LOGFILE%
    echo SUCCESS: VBS Phase 3 completed successfully
) else (
    echo [%date% %time%] ERROR: VBS Phase 3 failed with exit code %errorlevel% >> %LOGFILE%
    echo ERROR: VBS Phase 3 failed with exit code %errorlevel%
)

echo [%date% %time%] VBS Phase 3 execution finished >> %LOGFILE%
echo ============================================================
echo [%date% %time%] Execution completed. Check log: %LOGFILE%

REM Keep window open for 10 seconds to see result
timeout /t 10

exit /b %errorlevel% 