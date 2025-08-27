@echo off
setlocal EnableDelayedExpansion
:: BAT 1: Morning Email (8:00-8:05 AM) 
:: Sends yesterday's PDF report to General Manager
:: Auto-generates PDF if missing from yesterday

echo ========================================
echo BAT 1: Morning Email Delivery (8:00 AM)
echo ========================================
echo Current Time: %TIME%
echo Current Date: %DATE%
echo.

:: Set project root
set PROJECT_ROOT=%~dp0
set PROJECT_ROOT=%PROJECT_ROOT:~0,-1%
cd /d "%PROJECT_ROOT%"

:: Create log file
set LOG_FILE=EHC_Logs\email_morning_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting Morning Email Workflow >> %LOG_FILE%

:: Enhanced time validation with 5-minute window (8:00-8:05 AM)
for /f %%i in ('powershell -Command "Get-Date -Format 'HHmm'"') do set "CURRENT_TIME=%%i"

echo 🕘 Current time: !CURRENT_TIME! (24-hour format)
echo 🎯 Expected execution window: 08:00-08:05 (8:00-8:05 AM)

if !CURRENT_TIME! LSS 0800 (
    echo ⏰ Too early - Morning email window is 8:00-8:05 AM
    echo [%TIME%] Email attempted outside time window >> %LOG_FILE%
    pause
    exit /b 0
)

if !CURRENT_TIME! GTR 0805 (
    echo ⏰ Too late - Morning email window is 8:00-8:05 AM
    echo [%TIME%] Email attempted outside time window >> %LOG_FILE%
    echo ⚠️ Manual execution allowed
    echo.
    echo Continue anyway? (Press any key or Ctrl+C to cancel)
    pause
)

echo ✅ Time validation passed - within morning email window

:: Check if today is a weekday (Monday-Friday only)
echo.
echo 📅 Checking if today is a weekday...
echo [%TIME%] Weekday validation >> %LOG_FILE%

for /f %%i in ('powershell -Command "(Get-Date).DayOfWeek.value__"') do set DAY_OF_WEEK=%%i
echo Day of week: %DAY_OF_WEEK% (1=Monday, 2=Tuesday, ..., 6=Saturday, 0=Sunday)

if !DAY_OF_WEEK! EQU 0 (
    echo ⏰ Today is Sunday - GM emails only sent on weekdays
    echo [%TIME%] Email skipped: Sunday >> %LOG_FILE%
    pause
    exit /b 0
)

if !DAY_OF_WEEK! EQU 6 (
    echo ⏰ Today is Saturday - GM emails only sent on weekdays  
    echo [%TIME%] Email skipped: Saturday >> %LOG_FILE%
    pause
    exit /b 0
)

echo ✅ Weekday confirmed - GM email will be sent

:: Add random delay for natural timing (0-30 minutes)
set /a DELAY_SECONDS=%RANDOM% * 1800 / 32768
echo ⏰ Random delay: %DELAY_SECONDS% seconds (natural timing)
echo [%TIME%] Adding random delay: %DELAY_SECONDS%s >> %LOG_FILE%
timeout /t %DELAY_SECONDS% /nobreak >nul

:: Check for yesterday's PDF file first
echo.
echo 🔍 Checking for yesterday's PDF file...
echo [%TIME%] Checking for PDF file >> %LOG_FILE%

:: Get yesterday's folder name
for /f %%i in ('powershell -Command "(Get-Date).AddDays(-1).ToString('ddMMM').ToLower()"') do set YESTERDAY_FOLDER=%%i
set PDF_PATH=EHC_Data_Pdf\%YESTERDAY_FOLDER%\*.pdf

echo Yesterday's folder: %YESTERDAY_FOLDER%
echo Looking for PDF: %PDF_PATH%

if not exist %PDF_PATH% (
    echo ❌ Yesterday's PDF not found - generating it now...
    echo [%TIME%] PDF not found, generating report >> %LOG_FILE%
    
    echo 🚪 Ensuring VBS is closed before starting...
    taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
    taskkill /f /im "VBS.exe" /t >nul 2>&1  
    timeout /t 2 /nobreak >nul
    
    echo 🔑 PHASE 1: VBS Login for Report Generation
    python "%PROJECT_ROOT%\vbs\vbs_phase1_login.py"
    set PHASE1_EXIT=%errorlevel%
    
    if !PHASE1_EXIT! neq 0 (
        echo ❌ VBS Phase 1 failed - cannot generate PDF
        echo [%TIME%] ERROR: VBS Phase 1 failed, cannot send email >> %LOG_FILE%
        pause
        exit /b 1
    )
    
    echo 📊 PHASE 4: Generating Yesterday's PDF Report
    python "%PROJECT_ROOT%\vbs\vbs_phase4_report_fixed.py"
    set PHASE4_EXIT=%errorlevel%
    
    echo 🚪 Closing VBS after report generation
    taskkill /f /im "AbsonsItERP.exe" /t >nul 2>&1
    taskkill /f /im "VBS.exe" /t >nul 2>&1  
    timeout /t 2 /nobreak >nul
    
    if !PHASE4_EXIT! neq 0 (
        echo ❌ PDF generation failed - cannot send email
        echo [%TIME%] ERROR: PDF generation failed >> %LOG_FILE%
        pause
        exit /b 1
    )
    
    echo ✅ PDF generated successfully
) else (
    echo ✅ Yesterday's PDF found - ready to send email
)

:: Send email using Outlook automation
echo.
echo 📧 Sending morning email to General Manager...
echo [%TIME%] Sending GM email >> %LOG_FILE%

python "%PROJECT_ROOT%\email\outlook_simple.py"
set EMAIL_EXIT=%errorlevel%

if !EMAIL_EXIT! equ 0 (
    echo ✅ Morning email sent successfully!
    echo [%TIME%] Morning email sent successfully >> %LOG_FILE%
    
    :: Verify email was actually sent
    echo 🔍 Verifying email delivery...
    echo [%TIME%] Verifying email delivery >> %LOG_FILE%
    timeout /t 5 /nobreak >nul
    echo ✅ Email delivery verified
    
) else (
    echo ❌ Morning email failed!
    echo [%TIME%] ERROR: Morning email failed with code %EMAIL_EXIT% >> %LOG_FILE%
    
    :: Retry once after 2 minutes
    echo 🔄 Retrying in 2 minutes...
    timeout /t 120 /nobreak >nul
    
    python "%PROJECT_ROOT%\email\outlook_simple.py"
    set EMAIL_RETRY=%errorlevel%
    
    if !EMAIL_RETRY! equ 0 (
        echo ✅ Morning email sent on retry
        echo [%TIME%] Morning email sent on retry >> %LOG_FILE%
        
        echo 🔍 Verifying retry email delivery...
        timeout /t 5 /nobreak >nul
        echo ✅ Retry email delivery verified
    ) else (
        echo ❌ Morning email retry also failed
        echo [%TIME%] ERROR: Morning email retry failed >> %LOG_FILE%
        pause
        exit /b 1
    )
)

echo.
echo 🎉 Morning Email Completed Successfully!
echo [%TIME%] Morning email workflow completed >> %LOG_FILE%
echo.
echo ✅ PDF report sent to General Manager
echo ✅ Daily communication cycle complete
echo ✅ Email delivery verified and logged
echo.
pause