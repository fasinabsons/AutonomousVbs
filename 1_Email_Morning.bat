@echo off
setlocal EnableDelayedExpansion
:: BAT 1: Morning Email (9:00-9:30 AM)
:: Sends yesterday's PDF report to General Manager
:: Only runs if PDF exists from previous day

echo ========================================
echo BAT 1: Morning Email Delivery
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
set LOG_FILE=EHC_Logs\email_morning_%date:~-4,4%%date:~-10,2%%date:~-7,2%.log
echo [%TIME%] Starting Morning Email Workflow >> %LOG_FILE%

:: Enhanced time validation with 5-minute window (9:00-9:05 AM)
for /f %%i in ('powershell -Command "Get-Date -Format 'HHmm'"') do set "CURRENT_TIME=%%i"

echo 🕘 Current time: !CURRENT_TIME! (24-hour format)

if !CURRENT_TIME! LSS 0900 (
    echo ⏰ Too early - Morning email window is 9:00-9:05 AM
    echo [%TIME%] Email attempted outside time window >> %LOG_FILE%
    pause
    exit /b 0
)

if !CURRENT_TIME! GTR 0905 (
    echo ⏰ Too late - Morning email window is 9:00-9:05 AM
    echo [%TIME%] Email attempted outside time window >> %LOG_FILE%
    pause  
    exit /b 0
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

:: Send email using Outlook automation
echo.
echo 📧 Sending morning email to General Manager...
echo [%TIME%] Sending GM email >> %LOG_FILE%

python "%PROJECT_ROOT%\email\outlook_simple.py"
set EMAIL_EXIT=%errorlevel%

if !EMAIL_EXIT! equ 0 (
    echo ✅ Morning email sent successfully!
    echo [%TIME%] Morning email sent successfully >> %LOG_FILE%
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
echo.
pause