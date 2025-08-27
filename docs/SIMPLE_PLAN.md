# 🚀 MoonFlower 365-Day Automation - Simple 4-Step Plan

## **What You Get**: One PowerShell file that runs 10+ Python scripts automatically for 365 days

### **Step 1: Copy & Install** (2 minutes)
```powershell
# Copy project folder to any Windows PC, install Python packages
pip install selenium pyautogui opencv-python pandas xlwt openpyxl pywin32 psutil requests chromedriver-autoinstaller Pillow numpy
```

### **Step 2: Configure Chrome** (1 minute) 
```
Settings → Downloads → Turn OFF "Ask where to save files" → Allow multiple downloads
```

### **Step 3: Install Task** (30 seconds)
```powershell
# PowerShell as Admin, run once to setup 365-day auto-start
.\install.ps1 -TaskName "MoonFlowerAutomation" -ForCurrentUser -AtLogon -Daily8AM
```

### **Step 4: Run & Forget** (365 days)
```powershell
# Test immediately, then it runs automatically forever
.\AutomationMaster.ps1
```

---

## **What It Does Daily:**
- **9:00 AM**: Download 4 CSV files
- **12:30 PM**: Download 4 more CSV files (8 total)
- **12:35 PM**: Merge 8 CSV → 1 Excel file  
- **12:40 PM**: Upload Excel via VBS (30min-3hrs)
- **4:00 PM**: Close VBS safely
- **5:01 PM**: Generate PDF report
- **Next day 8:30 AM**: Email PDF to GM (weekdays only)

## **Built-in Intelligence:**
✅ **Weekday Rules**: Sunday PDF → Monday email, Monday PDF → Tuesday email  
✅ **Catch-up**: PC restart at 10 AM? Downloads CSVs → merges → uploads automatically  
✅ **File Validation**: Checks 4 CSV files, 8 CSV files, Excel exists, PDF exists  
✅ **VBS Management**: Auto-close at 4 PM, fresh login for reports, handles crashes  
✅ **Retry Logic**: 3 tries for CSV, 2 for Excel, 2 for reports with fresh sessions  
✅ **Error Emails**: Developer gets notified of failures automatically

## **Why PowerShell Over EXE:**
- **Works on any PC**: Copy-paste and run, no installation
- **Easy to modify**: Change times/emails without recompiling  
- **No Unicode errors**: Runs Python files directly
- **Better debugging**: Clear logs and error messages
- **Windows native**: Perfect Task Scheduler integration

## **File Structure** (all in one folder):
```
MoonFlower/
├── AutomationMaster.ps1    ← Main file (runs everything)
├── install.ps1             ← Setup script  
├── wifi/csv_downloader_resilient.py
├── excel/excel_generator.py
├── vbs/vbs_phase1_login.py (+ phases 2,3,4)
├── email/outlook_automation.py
├── utils/close_vbs.py + file_checker.py
└── Images/ (VBS screenshots for automation)
```

---

**Bottom Line**: One PowerShell script orchestrates all your Python files with perfect timing, error handling, and 365-day reliability. Just copy, install, and forget.
