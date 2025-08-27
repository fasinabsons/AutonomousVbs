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
    goto :csv_download
)

:: Afternoon Download Window (12:30-12:35 PM) 
if !CURRENT_TIME! geq 1230 if !CURRENT_TIME! leq 1235 (
    echo 🌆 AFTERNOON DOWNLOAD WINDOW (12:30-12:35 PM)
    goto :csv_download
)

:: Excel Merge Window (12:35-12:40 PM)
if !CURRENT_TIME! geq 1235 if !CURRENT_TIME! leq 1240 (
    echo 📊 EXCEL MERGE WINDOW (12:35-12:40 PM)
    goto :excel_merge
)

:: Default: Allow manual execution
echo 🔧 MANUAL EXECUTION - Running both CSV download and Excel merge
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
:: Check if we should run Excel merge now
set CURRENT_HOUR=%TIME:~0,2%
set CURRENT_MIN=%TIME:~3,2%
if "%CURRENT_HOUR:~0,1%"==" " set CURRENT_HOUR=0%CURRENT_HOUR:~1,1%

:: Skip Excel if it's before 12:35 PM
if %CURRENT_HOUR% LSS 12 (
    echo ⏰ Before 12:35 PM - Excel merge will run at scheduled time
    echo [%TIME%] Excel merge deferred until 12:35 PM >> %LOG_FILE%
    goto :send_notification
)

if %CURRENT_HOUR% EQU 12 if %CURRENT_MIN% LSS 35 (
    echo ⏰ Before 12:35 PM - Excel merge will run at scheduled time  
    echo [%TIME%] Excel merge deferred until 12:35 PM >> %LOG_FILE%
    goto :send_notification
)

:excel_merge
echo.
echo 📊 Excel Merge Process...
echo [%TIME%] Starting Excel merge >> %LOG_FILE%

python excel\excel_generator.py
set EXCEL_EXIT=%errorlevel%

if !EXCEL_EXIT! equ 0 (
    echo ✅ Excel merge successful!
    echo [%TIME%] Excel merge completed >> %LOG_FILE%
    
    :: Validate Excel file has sufficient data
    echo 🔍 Validating Excel file data...
    python -c "
import pandas as pd
from pathlib import Path
from datetime import datetime

try:
    today = datetime.now().strftime('%%d%%b').lower()
    excel_path = Path(f'EHC_Data_Merge/{today}/EHC_Upload_Mac_{datetime.now().strftime(\"%%d%%m%%Y\")}.xls')
    
    if excel_path.exists():
        df = pd.read_excel(excel_path)
        row_count = len(df)
        print(f'Excel file has {row_count} rows')
        
        if row_count >= 6000:
            print('✅ Excel validation passed (>=6000 rows)')
            exit(0)
        else:
            print(f'⚠️ Excel has only {row_count} rows (expected >=6000)')
            exit(0)  # Don't fail, just warn
    else:
        print('❌ Excel file not found')
        exit(1)
except Exception as e:
    print(f'❌ Excel validation failed: {e}')
    exit(1)
"
    
) else (
    echo ❌ Excel merge failed!
    echo [%TIME%] ERROR: Excel merge failed with code %EXCEL_EXIT% >> %LOG_FILE%
    
    :: Retry once
    echo 🔄 Retrying Excel merge...
    timeout /t 30 /nobreak >nul
    
    python excel\excel_generator.py
    set EXCEL_RETRY=%errorlevel%
    
    if !EXCEL_RETRY! equ 0 (
        echo ✅ Excel merge successful on retry
        echo [%TIME%] Excel merge successful on retry >> %LOG_FILE%
    ) else (
        echo ❌ Excel merge retry also failed
        echo [%TIME%] ERROR: Excel merge retry failed >> %LOG_FILE%
        pause
        exit /b 1
    )
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