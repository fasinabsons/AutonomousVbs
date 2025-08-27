# MoonFlower 365-Day Automation - Deployment PRD

## 🎯 OBJECTIVE
Create a simple, robust deployment solution that runs Python automation scripts continuously for 365 days on any Windows PC.

## 📋 REQUIREMENTS SUMMARY

### **Core Workflow**
1. **CSV Downloads**: 9:00 AM (4 files) + 12:30 PM (4 files) = 8 total files per day
2. **Excel Merge**: 12:35 PM after CSV completion (combines 8 CSV → 1 XLS file)  
3. **VBS Upload**: 12:40 PM (Phases 1-3: login → navigate → upload, takes 30min-3hrs)
4. **VBS Close**: 4:00 PM daily (force close VBS software)
5. **VBS Report**: 5:01 PM (fresh login → generate PDF report)
6. **GM Email**: Next day 8:30-9:30 AM (weekdays only, business day rules)
7. **Error Notifications**: Real-time via email to developer

### **Business Day Email Rules**
- **Monday**: Send Sunday's PDF (yesterday)
- **Tuesday**: Send Monday's PDF (yesterday)  
- **Wed-Fri**: Send previous day's PDF
- **Weekends**: No emails sent

### **Recovery & Resilience**
- **Catch-up Logic**: If PC starts at 10 AM with no CSVs → download twice → merge → upload → report
- **File Validation**: Check CSV count (≥4, ≥8), Excel exists, PDF exists after report
- **VBS Management**: Auto-close at 4 PM, fresh login for reports, handle "Not Responding" states
- **Retry Mechanisms**: 3 attempts for CSV, 2 for Excel, 2 for reports with fresh logins
- **Yesterday Fallback**: If today's PDF missing for email, try yesterday's report

## 🛠️ TECHNICAL SOLUTION: PowerShell Orchestrator

### **Why PowerShell Over EXE?**
1. **Portable**: Works from any folder, no installation required
2. **Maintainable**: Edit timing/slots without recompiling  
3. **Reliable**: Windows native, no codec issues or embedding problems
4. **Debuggable**: Clear logs, direct Python script execution
5. **Schedulable**: Native Task Scheduler integration

### **File Structure**
```
MoonFlower/
├── AutomationMaster.ps1        # Main orchestrator (24/7 loop)
├── install.ps1                 # Task Scheduler setup
├── uninstall.ps1              # Clean removal (keeps data)
├── wifi/
│   └── csv_downloader_resilient.py
├── excel/
│   └── excel_generator.py     # Uses xlwt for .xls format
├── vbs/
│   ├── vbs_phase1_login.py
│   ├── vbs_phase2_navigation_fixed.py
│   ├── vbs_phase3_upload_complete.py
│   └── vbs_phase4_report_fixed.py
├── email/
│   ├── outlook_automation.py   # GM emails (actual working file)
│   └── email_delivery.py      # Error notifications
├── utils/
│   ├── close_vbs.py           # Standalone VBS closer
│   ├── file_checker.py        # File validation & reports
│   ├── path_manager.py
│   ├── log_manager.py
│   └── file_manager.py
├── daily_folder_creator.py    # Midnight folder creation
└── Images/                    # VBS automation screenshots
    ├── phase2/
    ├── phase3/
    └── phase4/
```

## 🔧 IMPLEMENTATION DETAILS

### **AutomationMaster.ps1 Features**
```powershell
# Configurable parameters (top of script)
param(
    [string]$PythonExe = "python",
    [string]$MorningEmailStart = "08:30",
    [string]$VbsCloseTime = "16:00",
    [int]$CsvMin1 = 4,
    [int]$CsvMin2 = 8
)

# Extensible CSV slots
$CsvSlots = @(
    @{ Name='Slot1'; Start='09:00'; End='09:10'; MinCount=$CsvMin1 },
    @{ Name='Slot2'; Start='12:30'; End='12:40'; MinCount=$CsvMin2 }
    # Add more slots here...
)
```

### **Core Loop Logic**
```powershell
while ($true) {
    # 1. Midnight folder creation (00:00-00:05)
    Invoke-MidnightFolderCreation
    
    # 2. Daily state management
    Ensure-DailyFolders
    $state = Load-State
    Reset-IfNewDay $state
    
    # 3. Catch-up recovery (after 10:00 AM)
    Ensure-CatchUpIfNeeded
    
    # 4. File validation (non-fatal)
    Invoke-Py 'utils/file_checker.py'
    
    # 5. VBS closure (16:00-16:05)
    if (4PM window) { Close-VBS }
    
    # 6. GM Email (weekdays, business day PDF logic)
    if (weekday && email window && PDF exists) {
        Invoke-Py 'email/outlook_automation.py'
    }
    
    # 7. CSV Downloads (multiple configurable slots)
    foreach ($slot in $CsvSlots) {
        if (in slot window && not completed) {
            Invoke-Py 'wifi/csv_downloader_resilient.py'
            Validate file count ≥ MinCount
        }
    }
    
    # 8. Excel Merge (after CSV completion)
    if (excel window && CSV files ≥ 8) {
        Invoke-Py 'excel/excel_generator.py'
        Validate .xls file exists
    }
    
    # 9. VBS Upload (after Excel exists)
    if (vbs window && excel done) {
        Invoke-Py 'vbs/vbs_phase1_login.py'
        Invoke-Py 'vbs/vbs_phase2_navigation_fixed.py' 
        Invoke-Py 'vbs/vbs_phase3_upload_complete.py'
    }
    
    # 10. VBS Report (after upload, with fresh session)
    if (report window && upload done) {
        Close-VBS  # Fresh session
        Invoke-Py 'vbs/vbs_phase1_login.py'
        Invoke-Py 'vbs/vbs_phase4_report_fixed.py'
        Close-VBS  # Prevent crashes
        Validate PDF exists
    }
    
    Start-Sleep -Seconds 30  # Check every 30 seconds
}
```

## 📦 DEPLOYMENT PROCESS

### **3-Step Installation**
```bash
# 1. Copy files to any Windows PC
Copy-Item MoonFlower/ C:\MoonFlower\ -Recurse

# 2. Install Python dependencies
pip install selenium pyautogui opencv-python pandas xlwt openpyxl pywin32 psutil requests chromedriver-autoinstaller Pillow numpy

# 3. Install scheduled task (PowerShell as Admin)
cd C:\MoonFlower
.\install.ps1 -TaskName "MoonFlowerAutomation" -ForCurrentUser -AtLogon -Daily8AM
```

### **Chrome Configuration (Critical)**
- **Downloads**: Turn OFF "Ask where to save"
- **Multiple files**: Allow automatic downloads  
- **Display**: 100% Windows scale, 100% Chrome zoom
- **Language**: English UI for consistent selectors

### **VBS Software Setup**
- Install .NET VBS application
- Ensure login credentials saved
- Test manual login/navigation once

### **Outlook Configuration**  
- Configure organizational email account
- Test manual email sending
- Trust center: Allow macro notifications

## 🔄 OPERATION & MONITORING

### **Daily Schedule**
```
00:00 - Midnight folder creation
08:30 - GM email window starts (weekdays)
09:00 - CSV Slot 1 (4 files)
12:30 - CSV Slot 2 (4 more files)
12:35 - Excel merge (8 CSV → 1 XLS)
12:40 - VBS upload starts (30min-3hrs)
16:00 - VBS force close
17:01 - VBS report generation (fresh session)
```

### **File Validation**
```powershell
# Check today's files
python utils/file_checker.py --check all

# Generate report  
python utils/file_checker.py --check report --out EHC_Logs/file_report.json

# Clean suspicious files
python utils/file_checker.py --check cleanup
```

### **Manual Operations**
```powershell
# Force close VBS
python utils/close_vbs.py --force

# Test individual components
python wifi/csv_downloader_resilient.py
python excel/excel_generator.py  
python email/outlook_automation.py
```

## 📧 EMAIL CONFIGURATION

### **GM Email (outlook_automation.py)**
```python
self.config = {
    "gm_recipient": "ramon.logan@absons.ae",  # Easy to change
    "sender_account": "mohamed.fasin@absons.ae",
    "signature": "Professional blue/bold format"
}
```

### **Error Notifications (email_delivery.py)**
- CSV download failures → Developer email
- Report generation failures → Developer email  
- System crashes → Developer email
- **Send once per failure type** to avoid spam

## 🚀 ADVANTAGES OF THIS SOLUTION

### **Vs EXE Approach**
✅ **No Unicode/charmap errors** (direct Python execution)  
✅ **No embedding complexity** (files run as-is)
✅ **Easy customization** (edit .ps1 parameters)  
✅ **Better debugging** (clear Python output)
✅ **Portable** (works from any folder)

### **Vs BAT Files**  
✅ **Better error handling** (try-catch blocks)
✅ **Complex timing logic** (business day rules)
✅ **JSON state management** (persistent across reboots)
✅ **Advanced process management** (VBS closure)

### **Vs Python-only Master**
✅ **Windows integration** (Task Scheduler, process control)
✅ **Robust process execution** (timeout, retry, cleanup)
✅ **System-level operations** (taskkill, service management)

## 🎯 SUCCESS METRICS

### **Reliability Targets**
- **Uptime**: 365 days continuous operation
- **File Success Rate**: >99% (CSV, Excel, PDF generation)  
- **Email Success Rate**: >99% (GM reports, error notifications)
- **Recovery Time**: <1 hour after PC restart
- **Manual Intervention**: <1 time per month

### **Monitoring Indicators**
```
✅ GREEN: All daily tasks completed, files validated
🟡 YELLOW: Minor delays, retries successful  
🔴 RED: Multiple failures, manual intervention needed
```

## 📞 FINAL RECOMMENDATIONS

### **Best Deployment Option**: PowerShell Orchestrator
1. **Simplicity**: 3-step installation process
2. **Reliability**: Native Windows integration, robust error handling
3. **Maintainability**: Edit timings/emails without recompilation
4. **Portability**: Copy-paste to any Windows PC and run
5. **Debuggability**: Clear logs, direct Python script output

### **Fallback for Complex Networks**: 
If PowerShell execution policies are restricted:
- Use the enhanced `MOONFLOWER_COMPLETE.bat` file
- Same logic but in batch file format
- Less elegant but works in restricted environments

---

**Bottom Line**: The PowerShell orchestrator provides the perfect balance of simplicity, reliability, and maintainability for 365-day operation on any Windows PC.
