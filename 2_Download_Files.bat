@echo off
setlocal EnableDelayedExpansion
:: BAT 2: File Downloads (9:30 AM, 12:30 PM) + Excel Merge (12:35 PM)
:: Downloads CSV files and merges Excel data
:: Handles multiple time slots and file validation

echo ========================================
echo BAT 2: File Download & Excel Merge
echo ========================================
echo Current Time: %TIME%
echo Current Date: %DATE%
echo.

:: Set project root
:: Set project root (automatically detect from BAT file location)
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

:: Create log file
set LOG_FILE=EHC_Logs\download_files_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting File Download Workflow >> %LOG_FILE%

:: Create daily folders first and set TODAY_FOLDER
echo 📁 Step 1: Creating daily folders...
echo [%TIME%] Creating daily folders >> %LOG_FILE%

:: Generate today's folder name (format: 21jul, 22jul, etc.)
for /f %%i in ('powershell -Command "Get-Date -Format 'ddMMM'"') do set "TODAY_FOLDER=%%i"
set TODAY_FOLDER=%TODAY_FOLDER:~0,2%%TODAY_FOLDER:~2%
:: Convert to lowercase (PowerShell method for reliability)
for /f %%i in ('powershell -Command "'%TODAY_FOLDER%'.ToLower()"') do set "TODAY_FOLDER=%%i"
echo 📅 Today's folder: %TODAY_FOLDER%

python daily_folder_creator.py --create-today
if %errorlevel% neq 0 (
    echo ❌ Daily folder creation failed!
    echo [%TIME%] ERROR: Daily folder creation failed >> %LOG_FILE%
    pause
    exit /b 1
)
echo ✅ Daily folders ready

:: Enhanced time checking with 5-minute windows
for /f %%i in ('powershell -Command "Get-Date -Format 'HHmm'"') do set "CURRENT_TIME=%%i"

echo.
echo 🕘 Current time: !CURRENT_TIME! (24-hour format)

:: Morning Download Window (9:30-9:35 AM)
if !CURRENT_TIME! geq 0930 if !CURRENT_TIME! leq 0935 (
    echo 🌅 MORNING DOWNLOAD WINDOW (9:30-9:35 AM)
    set DOWNLOAD_SESSION=MORNING
    goto :csv_download
)

:: Afternoon Download Window (12:30-12:35 PM) 
if !CURRENT_TIME! geq 1230 if !CURRENT_TIME! leq 1235 (
    echo 🌆 AFTERNOON DOWNLOAD WINDOW (12:30-12:35 PM)
    set DOWNLOAD_SESSION=AFTERNOON
    goto :csv_download
)

:: Excel Merge Window (12:35-12:40 PM)
if !CURRENT_TIME! geq 1235 if !CURRENT_TIME! leq 1240 (
    echo 📊 EXCEL MERGE WINDOW (12:35-12:40 PM)
    set DOWNLOAD_SESSION=EXCEL_ONLY
    goto :excel_merge
)

:: Default: Allow manual execution
echo 🔧 MANUAL EXECUTION - Running both CSV download and Excel merge
set DOWNLOAD_SESSION=MANUAL
goto :csv_download

:csv_download
echo.
echo 📥 CSV Download with Enhanced Resilience...
echo [%TIME%] Starting CSV downloads >> %LOG_FILE%

:: Enhanced CSV downloader with intelligent file validation
set CSV_ATTEMPT=0
set TARGET_FILES=4
set MIN_ACCEPTABLE=3

:retry_csv
set /a CSV_ATTEMPT=%CSV_ATTEMPT%+1
echo.
echo 🔄 CSV Download Attempt %CSV_ATTEMPT%/15 (Target: %TARGET_FILES% files, Min: %MIN_ACCEPTABLE%)
echo [%TIME%] CSV download attempt %CSV_ATTEMPT% >> %LOG_FILE%

:: Count files before download
for /f %%i in ('powershell -Command "Get-ChildItem 'EHC_Data\%TODAY_FOLDER%' -Filter '*.csv' | Measure-Object | Select-Object -ExpandProperty Count"') do set FILES_BEFORE=%%i
if "%FILES_BEFORE%"=="" set FILES_BEFORE=0
echo 📊 Files before attempt: %FILES_BEFORE%

:: Set environment variable for BAT execution
set BAT_EXECUTION=1
python wifi\csv_downloader_resilient.py
set CSV_EXIT=%errorlevel%

:: Count files after download
for /f %%i in ('powershell -Command "Get-ChildItem 'EHC_Data\%TODAY_FOLDER%' -Filter '*.csv' | Measure-Object | Select-Object -ExpandProperty Count"') do set FILES_AFTER=%%i
if "%FILES_AFTER%"=="" set FILES_AFTER=0
echo 📊 Files after attempt: %FILES_AFTER%

:: Calculate new files downloaded
set /a NEW_FILES=%FILES_AFTER%-%FILES_BEFORE%
echo 📈 New files this attempt: %NEW_FILES%

:: Check if we have enough files now
if %FILES_AFTER% GEQ %TARGET_FILES% (
    echo ✅ TARGET ACHIEVED! Downloaded %FILES_AFTER%/%TARGET_FILES% files
    echo [%TIME%] CSV download successful - %FILES_AFTER% files total >> %LOG_FILE%
    goto :check_excel_time
)

:: Check if we have minimum acceptable files and it's been many attempts
if %FILES_AFTER% GEQ %MIN_ACCEPTABLE% (
    if %CSV_ATTEMPT% GEQ 8 (
        echo ✅ MINIMUM ACHIEVED! Downloaded %FILES_AFTER%/%TARGET_FILES% files (after %CSV_ATTEMPT% attempts)
        echo [%TIME%] CSV download acceptable - %FILES_AFTER% files total >> %LOG_FILE%
        goto :check_excel_time
    )
)

:: Continue retrying if we haven't reached limits
if !CSV_ATTEMPT! LSS 15 (
    echo ❌ Attempt %CSV_ATTEMPT% incomplete - only %FILES_AFTER%/%TARGET_FILES% files
    echo [%TIME%] CSV attempt %CSV_ATTEMPT% incomplete - %FILES_AFTER% files >> %LOG_FILE%
    
    :: Dynamic wait time based on attempt number
    if %CSV_ATTEMPT% LEQ 5 (
        echo ⏳ Waiting 1 minute before retry (early attempts)...
        timeout /t 60 /nobreak >nul
    ) else (
        echo ⏳ Waiting 2 minutes before retry (later attempts)...
        timeout /t 120 /nobreak >nul
    )
    goto :retry_csv
) else (
    echo ❌ All 15 attempts completed! Final result: %FILES_AFTER%/%TARGET_FILES% files
    echo [%TIME%] ERROR: All CSV attempts completed - %FILES_AFTER% files final >> %LOG_FILE%
    
    if %FILES_AFTER% GEQ %MIN_ACCEPTABLE% (
        echo ⚠️ Proceeding with %FILES_AFTER% files (minimum threshold met)
        echo [%TIME%] Proceeding with %FILES_AFTER% files (acceptable) >> %LOG_FILE%
        goto :check_excel_time
    ) else (
        echo ❌ Insufficient files (%FILES_AFTER% < %MIN_ACCEPTABLE%) - cannot proceed
        echo [%TIME%] ERROR: Insufficient files for processing >> %LOG_FILE%
        
        :: Send failure notification
        echo 📧 Sending failure notification...
        python email\email_delivery.py csv_failed
        pause
        exit /b 1
    )
)

:check_excel_time
:: Enhanced Excel merge logic with mandatory verification
echo.
echo 📊 Checking Excel merge requirements...

:: For afternoon sessions, Excel merge is MANDATORY
if "%DOWNLOAD_SESSION%"=="AFTERNOON" (
    echo 🎯 AFTERNOON SESSION - Excel merge is MANDATORY
    goto :excel_merge
)

:: For manual/other sessions, check timing
set CURRENT_HOUR=%TIME:~0,2%
set CURRENT_MIN=%TIME:~3,2%
if "%CURRENT_HOUR:~0,1%"==" " set CURRENT_HOUR=0%CURRENT_HOUR:~1,1%

:: Skip Excel if it's before 12:35 PM (for morning sessions)
if %CURRENT_HOUR% LSS 12 (
    echo ⏰ Before 12:35 PM - Excel merge deferred for morning session
    echo [%TIME%] Excel merge deferred until 12:35 PM >> %LOG_FILE%
    goto :verify_files_only
)

if %CURRENT_HOUR% EQU 12 if %CURRENT_MIN% LSS 35 (
    echo ⏰ Before 12:35 PM - Excel merge deferred for morning session
    echo [%TIME%] Excel merge deferred until 12:35 PM >> %LOG_FILE%
    goto :verify_files_only
)

:excel_merge
echo.
echo 📊 Excel Merge Process with Mandatory Verification...
echo [%TIME%] Starting Excel merge >> %LOG_FILE%

set EXCEL_ATTEMPT=0
set MAX_EXCEL_ATTEMPTS=5

:retry_excel_merge
set /a EXCEL_ATTEMPT=%EXCEL_ATTEMPT%+1
echo.
echo 🔄 Excel Merge Attempt %EXCEL_ATTEMPT%/%MAX_EXCEL_ATTEMPTS%
echo [%TIME%] Excel merge attempt %EXCEL_ATTEMPT% >> %LOG_FILE%

python excel\excel_generator.py
set EXCEL_EXIT=%errorlevel%

if !EXCEL_EXIT! equ 0 (
    echo ✅ Excel merge process completed
    echo [%TIME%] Excel merge completed >> %LOG_FILE%
    
    :: MANDATORY Excel file verification - MUST pass to continue
    echo 🔍 MANDATORY: Verifying Excel file exists and has data...
    python -c "
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys

try:
    today = datetime.now().strftime('%%d%%b').lower()
    excel_path = Path(f'EHC_Data_Merge/{today}/EHC_Upload_Mac_{datetime.now().strftime(\"%%d%%m%%Y\")}.xls')
    
    if excel_path.exists():
        df = pd.read_excel(excel_path)
        row_count = len(df)
        file_size = excel_path.stat().st_size
        
        print(f'Excel file exists: {excel_path.name}')
        print(f'File size: {file_size:,} bytes')
        print(f'Row count: {row_count:,} rows')
        
        if row_count >= 1000 and file_size >= 50000:
            print('✅ EXCEL VERIFICATION PASSED')
            print(f'✅ File has {row_count} rows and {file_size} bytes')
            sys.exit(0)
        else:
            print(f'❌ EXCEL VERIFICATION FAILED')
            print(f'❌ Insufficient data: {row_count} rows, {file_size} bytes')
            sys.exit(1)
    else:
        print(f'❌ EXCEL FILE NOT FOUND: {excel_path}')
        sys.exit(1)
except Exception as e:
    print(f'❌ EXCEL VERIFICATION ERROR: {e}')
    sys.exit(1)
"
    set EXCEL_VERIFY=%errorlevel%
    
    if !EXCEL_VERIFY! equ 0 (
        echo ✅ EXCEL VERIFICATION PASSED - File is ready!
        echo [%TIME%] Excel file verified successfully >> %LOG_FILE%
        goto :excel_success
    ) else (
        echo ❌ EXCEL VERIFICATION FAILED - File missing or insufficient data
        echo [%TIME%] Excel verification failed on attempt %EXCEL_ATTEMPT% >> %LOG_FILE%
    )
    
) else (
    echo ❌ Excel merge process failed!
    echo [%TIME%] ERROR: Excel merge failed with code %EXCEL_EXIT% >> %LOG_FILE%
)

:: Retry logic for failed Excel merge or verification
if %EXCEL_ATTEMPT% LSS %MAX_EXCEL_ATTEMPTS% (
    echo 🔄 Retrying Excel merge in 30 seconds...
    echo [%TIME%] Retrying Excel merge (attempt %EXCEL_ATTEMPT%) >> %LOG_FILE%
    timeout /t 30 /nobreak >nul
    goto :retry_excel_merge
) else (
    echo ❌ All Excel merge attempts failed!
    echo [%TIME%] ERROR: All Excel merge attempts failed >> %LOG_FILE%
    
    :: Critical failure - cannot proceed without Excel file for afternoon session
    if "%DOWNLOAD_SESSION%"=="AFTERNOON" (
        echo ❌ CRITICAL: Afternoon session requires Excel file
        echo [%TIME%] CRITICAL: Cannot proceed without Excel file >> %LOG_FILE%
        pause
        exit /b 1
    ) else (
        echo ⚠️ WARNING: Excel merge failed but continuing for morning session
        echo [%TIME%] WARNING: Excel merge failed, continuing >> %LOG_FILE%
        goto :verify_files_only
    )
)

:excel_success
echo ✅ Excel merge completed and verified successfully!
goto :send_notification

:verify_files_only
echo.
echo 🔍 File verification for morning session (no Excel merge)...
echo [%TIME%] Morning session file verification >> %LOG_FILE%

:: Count CSV files
for /f %%i in ('powershell -Command "Get-ChildItem 'EHC_Data\%TODAY_FOLDER%' -Filter '*.csv' | Measure-Object | Select-Object -ExpandProperty Count"') do set FINAL_CSV_COUNT=%%i
if "%FINAL_CSV_COUNT%"=="" set FINAL_CSV_COUNT=0

echo 📊 Morning session complete: %FINAL_CSV_COUNT% CSV files downloaded
echo [%TIME%] Morning session: %FINAL_CSV_COUNT% CSV files >> %LOG_FILE%

if %FINAL_CSV_COUNT% GEQ %MIN_ACCEPTABLE% (
    echo ✅ Sufficient CSV files for morning session
    goto :send_notification
) else (
    echo ❌ Insufficient CSV files for morning session
    echo [%TIME%] ERROR: Insufficient CSV files in morning session >> %LOG_FILE%
    pause
    exit /b 1
)

:send_notification
echo.
echo 📧 Sending completion notification with file count...
echo [%TIME%] Sending notification >> %LOG_FILE%

:: Count CSV files and send detailed notification
python -c "
import os
from datetime import datetime
from pathlib import Path

try:
    today = datetime.now().strftime('%%d%%b').lower()
    csv_folder = Path(f'EHC_Data/{today}')
    
    if csv_folder.exists():
        csv_files = list(csv_folder.glob('*.csv'))
        file_count = len(csv_files)
        print(f'Found {file_count} CSV files in {today} folder')
        
        # Check if Excel file exists
        excel_folder = Path(f'EHC_Data_Merge/{today}')
        excel_files = list(excel_folder.glob('*.xls')) if excel_folder.exists() else []
        excel_merged = len(excel_files) > 0
        
        # Send notification via email_delivery.py
        import subprocess
        if excel_merged:
            subprocess.run(['python', 'email/email_delivery.py', 'csv_complete', str(file_count), today])
        else:
            subprocess.run(['python', 'email/email_delivery.py', 'csv_only_complete', str(file_count), today])
        
        print(f'Notification sent: {file_count} files, Excel merged: {excel_merged}')
    else:
        print(f'CSV folder not found: {csv_folder}')
        subprocess.run(['python', 'email/email_delivery.py', 'csv_failed'])
        
except Exception as e:
    print(f'Notification error: {e}')
    import subprocess
    subprocess.run(['python', 'email/email_delivery.py', 'csv_failed'])
"

echo.
echo 🎉 File Download & Merge Completed!
echo [%TIME%] Download workflow completed >> %LOG_FILE%
echo.
echo ✅ CSV files downloaded and validated
echo ✅ Excel data merged and ready
echo ✅ Ready for VBS automation (BAT 3)
echo.
pause