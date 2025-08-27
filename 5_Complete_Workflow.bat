@echo off
:: BAT 5: Complete Workflow (All Steps)
:: Runs entire automation from start to finish
:: For manual execution or complete automation testing

echo ================================================================
echo BAT 5: Complete MoonFlower Automation Workflow
echo ================================================================
echo Current Time: %TIME%
echo Current Date: %DATE%
echo.
echo 🤖 COMPLETE AUTOMATION SEQUENCE:
echo   1. Morning Email (if yesterday's PDF exists)
echo   2. File Downloads + Excel Merge
echo   3. VBS Upload (Phase 1 → 2 → 3)
echo   4. VBS Report (Phase 1 → 4)
echo   5. Final Status Report
echo.

:: Set project root
:: Set project root (automatically detect from BAT file location)
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

:: Create log file
set LOG_FILE=EHC_Logs\complete_workflow_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting Complete Workflow >> %LOG_FILE%

:: Ask user for confirmation
set /p CONFIRM="Continue with complete automation? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo User cancelled automation
    exit /b 0
)

echo.
echo 📧 STEP 1: Morning Email
echo ================================================================
echo [%TIME%] Step 1: Morning Email >> %LOG_FILE%

:: Check if we should send morning email
set HOUR=%TIME:~0,2%
if "%HOUR:~0,1%"==" " set HOUR=0%HOUR:~1,1%

if %HOUR% GEQ 9 if %HOUR% LEQ 10 (
    echo 🕘 Within morning email window - sending email
    call 1_Email_Morning.bat
    if %errorlevel% neq 0 (
        echo ⚠️ Morning email failed, but continuing...
        echo [%TIME%] WARNING: Morning email failed >> %LOG_FILE%
    )
) else (
    echo ⏰ Outside morning email window (9-10 AM) - skipping
    echo [%TIME%] Morning email skipped (outside time window) >> %LOG_FILE%
)

echo.
echo 📥 STEP 2: File Downloads & Excel Merge
echo ================================================================
echo [%TIME%] Step 2: File Downloads >> %LOG_FILE%

call 2_Download_Files.bat
if %errorlevel% neq 0 (
    echo ❌ File downloads failed!
    echo [%TIME%] ERROR: File downloads failed >> %LOG_FILE%
    goto :workflow_error
)
echo ✅ Step 2 completed successfully

echo.
echo ⬆️ STEP 3: VBS Upload Process
echo ================================================================
echo [%TIME%] Step 3: VBS Upload >> %LOG_FILE%

call 3_VBS_Upload.bat
if %errorlevel% neq 0 (
    echo ❌ VBS upload failed!
    echo [%TIME%] ERROR: VBS upload failed >> %LOG_FILE%
    goto :workflow_error
)
echo ✅ Step 3 completed successfully

echo.
echo 📊 STEP 4: VBS Report Generation
echo ================================================================
echo [%TIME%] Step 4: VBS Report >> %LOG_FILE%

call 4_VBS_Report.bat
if %errorlevel% neq 0 (
    echo ❌ VBS report failed!
    echo [%TIME%] ERROR: VBS report failed >> %LOG_FILE%
    goto :workflow_error
)
echo ✅ Step 4 completed successfully

echo.
echo 📋 STEP 5: Final Status Report
echo ================================================================
echo [%TIME%] Step 5: Final Status Report >> %LOG_FILE%

:: Verify all components
echo 🔍 Verifying automation results...

:: Check CSV files - Get today's folder in ddMMM format (e.g., 31jul)
for /f %%i in ('powershell -Command "(Get-Date).ToString('ddMMM').ToLower()"') do set TODAY_FOLDER=%%i

echo.
echo 📊 AUTOMATION RESULTS:
echo ----------------------------------------------------------------

:: CSV Files
set CSV_COUNT=0
for %%f in (EHC_Data\%TODAY_FOLDER%\*.csv) do set /a CSV_COUNT+=1
if %CSV_COUNT% GEQ 4 (
    echo ✅ CSV Files: %CSV_COUNT% files downloaded
) else (
    echo ⚠️ CSV Files: Only %CSV_COUNT% files (expected 4+)
)

:: Excel File
if exist EHC_Data_Merge\%TODAY_FOLDER%\*.xls (
    echo ✅ Excel File: Merged successfully
) else (
    echo ❌ Excel File: Not found
)

:: PDF File
if exist EHC_Data_Pdf\%TODAY_FOLDER%\*.pdf (
    echo ✅ PDF Report: Generated successfully
    for %%f in (EHC_Data_Pdf\%TODAY_FOLDER%\*.pdf) do (
        echo    📄 File: %%~nxf (%%~zf bytes)
    )
) else (
    echo ❌ PDF Report: Not found
)

echo.
echo 📧 Sending final completion notification...
python email\email_delivery.py workflow_complete
if %errorlevel% neq 0 (
    echo ⚠️ Final notification failed
    echo [%TIME%] WARNING: Final notification failed >> %LOG_FILE%
)

echo.
echo 🎉 COMPLETE WORKFLOW FINISHED SUCCESSFULLY!
echo ================================================================
echo [%TIME%] Complete workflow finished successfully >> %LOG_FILE%
echo.
echo ✅ All 4 automation steps completed
echo ✅ Data processed and ready
echo ✅ PDF generated for tomorrow's email
echo ✅ System ready for next cycle
echo.
echo 📅 Next automated actions:
echo   • Tomorrow 9:00-9:30 AM: Email PDF to GM
echo   • Tomorrow 9:30 AM: Download new CSV files
echo   • Continue daily automation cycle
echo.
goto :workflow_end

:workflow_error
echo.
echo ❌ WORKFLOW ERROR DETECTED
echo ================================================================
echo [%TIME%] Workflow error detected >> %LOG_FILE%
echo.
echo 💡 Troubleshooting Steps:
echo   1. Check individual BAT file logs
echo   2. Verify VBS application is running
echo   3. Check network connectivity
echo   4. Validate Excel/PDF file permissions
echo.
echo 📧 Sending error notification...
python email\email_delivery.py workflow_error
pause
exit /b 1

:workflow_end
echo Press any key to exit...
pause >nul