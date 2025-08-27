# MoonFlower Automation - Complete PC Setup Guide
## 365-Day Continuous Operation Setup

This guide will help you set up the MoonFlower automation system on any new PC to run continuously for 365 days without issues.

---

## 📋 **REQUIREMENTS CHECKLIST**

### Hardware Requirements
- [ ] Windows 10/11 PC
- [ ] 8GB RAM minimum (16GB recommended)
- [ ] 100GB free disk space
- [ ] Stable internet connection
- [ ] PC will stay powered on 24/7

### Software Requirements
- [ ] Python 3.9 or higher
- [ ] Google Chrome browser
- [ ] Microsoft Outlook (configured with mohamed.fasin@absons.ae)
- [ ] VBS application installed and configured

---

## 🚀 **STEP 1: SYSTEM PREPARATION**

### 1.1 Windows Settings
```batch
# Disable Windows Updates restart
# Go to: Settings > Update & Security > Windows Update > Advanced options
# Set "Restart this device as soon as possible" to OFF

# Disable sleep/hibernation
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change hibernate-timeout-ac 0
powercfg /change hibernate-timeout-dc 0

# Set power plan to High Performance
powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
```

### 1.2 Create Automation User (Optional but Recommended)
```batch
# Create dedicated user for automation
net user MoonFlowerBot YourPassword123! /add
net localgroup administrators MoonFlowerBot /add
```

### 1.3 Disable Screen Lock
```batch
# Registry edit to disable lock screen
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization" /v NoLockScreen /t REG_DWORD /d 1 /f
```

---

## 🐍 **STEP 2: PYTHON INSTALLATION**

### 2.1 Download and Install Python
1. Download Python 3.11 from https://python.org
2. **IMPORTANT**: Check "Add Python to PATH" during installation
3. Verify installation:
```cmd
python --version
pip --version
```

### 2.2 Install Required Packages
```cmd
# Navigate to project directory
cd C:\MoonFlowerAutomation

# Install all packages from requirements.txt
pip install -r requirements.txt

# Manual installation if requirements.txt fails
pip install pyautogui==0.9.54
pip install opencv-python==4.8.1.78
pip install pillow==10.0.1
pip install pywin32==306
pip install selenium==4.15.2
pip install pandas==2.1.4
pip install openpyxl==3.1.2
pip install xlwt==1.3.0
pip install python-dateutil==2.8.2
```

---

## 📁 **STEP 3: PROJECT SETUP**

### 3.1 Create Project Directory
```cmd
mkdir C:\MoonFlowerAutomation
cd C:\MoonFlowerAutomation
```

### 3.2 Copy Project Files
Copy the complete Automata2 folder contents to `C:\MoonFlowerAutomation\`

### 3.3 Update Project Paths
Edit these files to match your PC:
- `COMPLETE_365DAY_AUTOMATION.bat` → Change `PROJECT_ROOT` path
- `SETUP_365DAY_SCHEDULER.bat` → Change `PROJECT_ROOT` path

### 3.4 Verify Project Structure
```
C:\MoonFlowerAutomation\
├── email/
│   ├── email_delivery.py        (Notifications to you)
│   └── outlook_automation.py    (GM emails - ONLY yesterday's PDF)
├── excel/
│   └── excel_generator.py       (CSV to Excel merge)
├── vbs/
│   ├── vbs_phase1_login.py
│   ├── vbs_phase2_navigation_fixed.py
│   ├── vbs_phase3_upload_calibrated.py  ← USES THIS VERSION (100% tested)
│   └── vbs_phase4_report_fixed.py
├── wifi/
│   └── csv_downloader_simple.py (CSV downloads)
├── Images/
│   └── phase3/ (all button images for VBS automation)
├── EHC_Data/               (CSV files by date)
├── EHC_Data_Merge/         (Excel files by date)
├── EHC_Data_Pdf/           (PDF reports by date)
├── EHC_Logs/               (All logs by date)
├── requirements.txt        (All Python packages)
├── COMPLETE_365DAY_AUTOMATION.bat  (MAIN automation file)
└── SETUP_365DAY_SCHEDULER.bat     (Windows Task Scheduler setup)
```

---

## 🔧 **STEP 4: AUTOMATION FILES SETUP**

### 4.1 Main Automation File
`COMPLETE_365DAY_AUTOMATION.bat` includes:
- ✅ **Daily folder creation**
- ✅ **CSV downloads** (wifi\csv_downloader_simple.py)
- ✅ **Excel merge** (excel\excel_generator.py)
- ✅ **VBS Phase 1** (login)
- ✅ **VBS Phase 2** (navigation)
- ✅ **VBS Phase 3** (upload) - **USES vbs_phase3_upload_calibrated.py** 
- ✅ **VBS Phase 4** (PDF generation)
- ✅ **Email notifications** (to you)
- ✅ **GM email** (ONLY if yesterday's PDF exists)

### 4.2 VBS Phase 3 Details
**File Used**: `vbs\vbs_phase3_upload_calibrated.py`
**Features**:
- ✅ 100% tested update button detection
- ✅ Calibrated confidence levels (0.8, 0.7, 0.6, 0.5, 0.4)
- ✅ Fast scanning (0.2 second intervals)
- ✅ Window maximization for button visibility
- ✅ 3-hour upload monitoring with popup detection
- ✅ Multiple update button image variants

### 4.3 Email Logic
**GM Email**: ONLY sends if yesterday's PDF exists
- ✅ Looks for yesterday's date folder only
- ✅ No email if no yesterday PDF
- ✅ Professional formatting with blue bold signature
- ✅ From: mohamed.fasin@absons.ae → To: ramon.logan@absons.ae

**Notifications**: Simple status updates to faseenm@gmail.com
- ✅ CSV complete, Excel complete, Upload complete, etc.

---

## 📧 **STEP 5: EMAIL CONFIGURATION**

### 5.1 Gmail App Password (For Notifications)
1. Go to Google Account settings for `fasin.absons@gmail.com`
2. Enable 2-factor authentication
3. Generate App Password 
4. Current password in system: `zrxj vfjt wjos wkwy`

### 5.2 Outlook Configuration (For GM Emails)
1. Install Microsoft Outlook
2. Add account: `mohamed.fasin@absons.ae`
3. Test email sending manually
4. Verify the account is set as default sender

---

## 🌐 **STEP 6: BROWSER SETUP**

### 6.1 Chrome Setup
```cmd
# Download ChromeDriver from https://chromedriver.chromium.org/
# Place chromedriver.exe in C:\MoonFlowerAutomation\ or in PATH
```

### 6.2 VBS Application
1. Install VBS application
2. Configure login credentials
3. Test manual login and navigation
4. All button images are in `Images/phase3/` folder

---

## ⚙️ **STEP 7: INSTALL 365-DAY AUTOMATION**

### 7.1 Run Scheduler Setup (AS ADMINISTRATOR)
```cmd
# Right-click and "Run as administrator"
cd C:\MoonFlowerAutomation
SETUP_365DAY_SCHEDULER.bat
```

### 7.2 Verify Scheduled Tasks
```cmd
# Check all tasks are created
schtasks /query /fo table | findstr MoonFlower
```

### 7.3 Scheduled Tasks Created
- ✅ **MoonFlower_Morning_Data**: 9:30 AM daily (CSV + Excel + VBS + Email)
- ✅ **MoonFlower_Afternoon_Data**: 12:30 PM daily (CSV + Excel + VBS + Email)  
- ✅ **MoonFlower_GM_Email**: 8:00 AM daily (GM email ONLY if yesterday PDF exists)

---

## 🔒 **STEP 8: SECURITY & PERSISTENCE**

### 8.1 Auto-Login Configuration (For 365-Day Operation)
```batch
# Enable auto-login for automation user
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v AutoAdminLogon /t REG_SZ /d 1 /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultUserName /t REG_SZ /d "MoonFlowerBot" /f
reg add "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon" /v DefaultPassword /t REG_SZ /d "YourPassword123!" /f
```

### 8.2 Startup Verification
Create `startup_check.bat` in Windows Startup folder:
```batch
@echo off
cd /d "C:\MoonFlowerAutomation"
echo System started at %TIME% on %DATE% >> startup.log
echo Checking Python installation...
python --version >> startup.log 2>&1
echo Checking project files...
dir *.bat >> startup.log 2>&1
```

---

## 📊 **STEP 9: MONITORING SETUP**

### 9.1 System Health Check
The automation creates logs in `EHC_Logs\{date}\` folders:
- ✅ CSV download logs
- ✅ Excel merge logs
- ✅ VBS automation logs (all phases)
- ✅ Email delivery logs

### 9.2 Email Monitoring
You'll receive notifications for:
- ✅ CSV downloads complete
- ✅ Excel merge complete
- ✅ VBS upload complete
- ✅ Automation cycle complete

---

## 🚨 **STEP 10: TROUBLESHOOTING**

### 10.1 Common Issues
```batch
# If tasks don't run:
1. Check Task Scheduler logs
2. Verify Python path in System PATH
3. Check file permissions (Run as Administrator)
4. Test BAT files manually first

# If VBS automation fails:
1. Check VBS application is installed
2. Verify Images/phase3/ folder has all button images
3. Test vbs_phase3_upload_calibrated.py manually
4. Check screen resolution matches training images

# If emails don't send:
1. Verify Outlook is configured with mohamed.fasin@absons.ae
2. Check internet connection
3. Test Gmail app password for notifications
4. Check if yesterday's PDF exists for GM email

# If GM email not sending:
1. Run VBS Phase 1 → Phase 4 to generate yesterday's PDF first
2. Then GM email will send automatically next day
```

### 10.2 Recovery Commands
```batch
# Restart Task Scheduler service
net stop "Task Scheduler"
net start "Task Scheduler"

# Reset automation schedule
cd C:\MoonFlowerAutomation
SETUP_365DAY_SCHEDULER.bat

# Test individual components
python wifi\csv_downloader_simple.py
python excel\excel_generator.py
python vbs\vbs_phase3_upload_calibrated.py
python email\outlook_automation.py
```

---

## ✅ **STEP 11: FINAL VERIFICATION**

### 11.1 Manual Test Run
```batch
cd C:\MoonFlowerAutomation

# Test complete automation
COMPLETE_365DAY_AUTOMATION.bat

# Test individual components
python wifi\csv_downloader_simple.py
python excel\excel_generator.py
python email\email_delivery.py automation_complete
python email\outlook_automation.py
```

### 11.2 Verification Checklist
- [ ] All Python packages installed (pip install -r requirements.txt)
- [ ] Scheduled tasks created and active
- [ ] VBS application responds to automation
- [ ] Email notifications work (to faseenm@gmail.com)
- [ ] GM email works (ONLY if yesterday PDF exists)
- [ ] System stays awake and connected
- [ ] Auto-login works after restart
- [ ] All folder structures exist with proper permissions

---

## 🎯 **STEP 12: GO LIVE**

### 12.1 Final Setup
1. Restart the PC to test auto-login
2. Verify all scheduled tasks run at correct times
3. Monitor for 24 hours to ensure stability
4. Set up remote monitoring if needed

### 12.2 365-Day Operation Schedule
- ✅ **9:30 AM**: CSV downloads + Excel merge + VBS automation + PDF generation
- ✅ **12:30 PM**: CSV downloads + Excel merge + VBS automation + PDF generation
- ✅ **8:00 AM**: GM email delivery (ONLY if yesterday's PDF exists)
- ✅ **Continuous**: Email notifications to you for all milestones

### 12.3 What Happens Each Day
1. **Morning (9:30 AM)**: Complete automation cycle
2. **Afternoon (12:30 PM)**: Complete automation cycle  
3. **Next Morning (8:00 AM)**: GM email with yesterday's PDF
4. **Continuous**: System creates daily folders, processes data, sends notifications

---

## 📞 **SUPPORT & MAINTENANCE**

### Daily Monitoring
- Check notification emails from faseenm@gmail.com
- Verify files are being created in dated folders
- Monitor disk space usage

### Weekly Tasks
- Check Task Scheduler for failed tasks
- Review automation logs for errors
- Clear old log files (keep last 7 days)

### Monthly Tasks
- Update system if needed
- Verify all credentials are still valid
- Test manual override procedures

---

## 🔧 **ONE-CLICK INSTALLATION SCRIPT**

Save as `INSTALL_MOONFLOWER.bat` and run as Administrator:

```batch
@echo off
echo Installing MoonFlower 365-Day Automation...

:: Create directory
mkdir C:\MoonFlowerAutomation 2>nul
cd /d C:\MoonFlowerAutomation

:: Copy project files (user must do this manually)
echo.
echo STEP 1: Please copy all Automata2 files to C:\MoonFlowerAutomation
echo Press any key when ready...
pause

:: Install Python packages
echo.
echo STEP 2: Installing Python packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Installing packages individually...
    pip install pyautogui opencv-python pillow pywin32 selenium pandas openpyxl
)

:: Set up scheduled tasks
echo.
echo STEP 3: Setting up scheduled tasks...
call SETUP_365DAY_SCHEDULER.bat

:: Configure system settings
echo.
echo STEP 4: Configuring system settings...
powercfg /change standby-timeout-ac 0
powercfg /change hibernate-timeout-ac 0

echo.
echo ================================================================
echo MoonFlower 365-Day Automation Installation Complete!
echo ================================================================
echo.
echo NEXT STEPS:
echo 1. Configure Outlook with mohamed.fasin@absons.ae
echo 2. Install and configure VBS application
echo 3. Test manually: COMPLETE_365DAY_AUTOMATION.bat
echo 4. Verify scheduled tasks: schtasks /query /fo table ^| findstr MoonFlower
echo.
echo The system will run automatically 365 days a year!
echo ================================================================
pause
```

---

## 🏁 **CONCLUSION**

Following this guide will set up a completely autonomous MoonFlower automation system that runs 365 days a year without user intervention. The system will:

- ✅ **Download CSV files** automatically (2x daily)
- ✅ **Merge Excel data** daily
- ✅ **Upload data via VBS** automation (uses calibrated Phase 3)
- ✅ **Generate PDF reports** with yesterday's date
- ✅ **Send GM emails** ONLY when yesterday's PDF exists
- ✅ **Provide notifications** on all milestones to you
- ✅ **Handle errors** and retry automatically
- ✅ **Keep detailed logs** for monitoring
- ✅ **Create daily folders** automatically
- ✅ **Work on any PC** with this setup guide

**The PC becomes a dedicated automation server that requires minimal maintenance while providing reliable daily operations for 365 days.** 