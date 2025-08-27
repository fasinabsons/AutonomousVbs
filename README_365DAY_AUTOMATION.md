# MoonFlower 365-Day Automation System

## 🎯 Complete Daily Schedule

This system runs **365 days without interruption**, handling all aspects of MoonFlower data collection, processing, and reporting.

### Daily Timeline

| Time | Task | Description |
|------|------|-------------|
| **9:30 AM** | 📥 Morning Download | Download CSV files from WiFi system |
| **12:30 PM** | 📊 Afternoon Download + Excel | Download more CSV files and merge into Excel |
| **1:00 PM** | ⬆️ VBS Upload | Start 3-hour upload process to VBS system |
| **1:58 PM** | 🔧 Pre-Restart Prep | Save automation state before restart |
| **2:00 PM** | 🔄 PC Restart | Restart PC for VBS reliability |
| **2:02 PM** | 🚀 Post-Restart Resume | Resume automation after restart |
| **4:00 PM** | 📋 Report Generation | Generate PDF reports in VBS |
| **8:00 PM** | 📧 Email Delivery | Send PDF report to General Manager |

## 🔐 Advanced Features

### Lock/Unlock State Handling
- ✅ Runs when PC is **locked or unlocked**
- ✅ Runs when **user is logged in or not**
- ✅ Runs with **highest system privileges**

### Power Management
- ✅ Runs on **battery or plugged in**
- ✅ **Wakes PC from sleep** if needed
- ✅ **Prevents sleep** during critical operations

### PC Restart Continuity
- 🔧 **Smart state saving** before 2:00 PM restart
- 🚀 **Automatic resumption** of interrupted uploads
- 📊 **Process tracking** across restarts

## 🚀 Installation

### Quick Install
```cmd
# Run as Administrator
INSTALL_365DAY_COMPLETE_SCHEDULE.bat
```

### Manual Install
```powershell
# Run as Administrator
.\Setup_365Day_Complete_Schedule.ps1 -Install
```

## 📊 Management

### Check Status
```cmd
CHECK_365DAY_SCHEDULE_STATUS.bat
```

### Uninstall
```cmd
UNINSTALL_365DAY_SCHEDULE.bat
```

## 📁 File Structure

### Core BAT Files
- `1_Email_Morning.bat` → **8:00 PM Email** (renamed for evening use)
- `2_Download_Files.bat` → **9:30 AM & 12:30 PM Downloads**
- `3_VBS_Upload.bat` → **1:00 PM Upload**
- `4_VBS_Report.bat` → **4:00 PM Reports**

### PowerShell Infrastructure
- `Setup_365Day_Complete_Schedule.ps1` → Main scheduler installer
- `scripts\handle_2pm_restart.ps1` → Restart continuity handler

### VBS Automation Scripts
- `vbs\vbs_phase1_login.py` → VBS login automation
- `vbs\vbs_phase2_navigation_fixed.py` → VBS navigation
- `vbs\vbs_phase3_upload_complete.py` → **Simplified upload logic**
- `vbs\vbs_phase4_report_fixed.py` → PDF report generation

### Support Systems
- `wifi\csv_downloader_resilient.py` → Enhanced CSV downloader
- `excel\excel_generator.py` → Excel merge system
- `email\outlook_simple.py` → Email automation
- `vbs\vbs_audio_detector.py` → Popup sound detection

## 🔧 Key Improvements Made

### 1. Email System (8:00 PM)
- ✅ **Auto-generates missing PDFs** before sending
- ✅ **Email delivery verification**
- ✅ **Handles yesterday's data** correctly
- ✅ **Proper time window validation**

### 2. Download System (9:30 AM & 12:30 PM)
- ✅ **Mandatory Excel verification** for afternoon sessions
- ✅ **File count validation** before proceeding
- ✅ **Retry logic** with progressive delays
- ✅ **Session-specific handling** (morning vs afternoon)

### 3. VBS Upload (1:00 PM)
- ✅ **Simplified Phase 3 logic**: Import → ENTER → Wait 5s → Click Update → 3h wait
- ✅ **Multiple update button variants** for reliability
- ✅ **Time window validation** (1:00-1:05 PM)
- ✅ **Excel dependency checking**

### 4. PC Restart (2:00 PM)
- ✅ **State preservation** before restart
- ✅ **Graceful VBS closure** with process tracking
- ✅ **Automatic resumption** after restart
- ✅ **Upload continuity** across restarts

### 5. Report Generation (4:00 PM)
- ✅ **Time window validation** (4:00-4:05 PM)
- ✅ **Clean VBS startup/shutdown**
- ✅ **PDF verification** after generation

## ⚠️ Important Notes

### System Requirements
- **Windows 10/11** with Task Scheduler
- **Administrator privileges** for installation
- **Python 3.x** with required packages
- **Microsoft Outlook** for email functionality
- **VBS application** access

### Critical Paths
```
⚠️ DO NOT MOVE THIS FOLDER after installation!
   Moving will break all scheduled tasks.
   
⚠️ Keep all BAT files in the root directory
   Scheduled tasks reference absolute paths.
```

### Network Dependencies
- **WiFi system** must be accessible for CSV downloads
- **VBS system** must be accessible for uploads
- **Outlook/Exchange** must be configured for emails

## 🔍 Troubleshooting

### Check Task Status
```cmd
CHECK_365DAY_SCHEDULE_STATUS.bat
```

### Common Issues

#### Tasks Not Running
1. Check **Administrator privileges**
2. Verify **Task Scheduler service** is running
3. Check **execution policy**: `Set-ExecutionPolicy RemoteSigned`

#### VBS Upload Issues
1. Check **VBS application** installation
2. Verify **network connectivity**
3. Check **image files** in `Images\phase3\` folder

#### Email Issues  
1. Verify **Outlook configuration**
2. Check **sender account**: mohamed.fasin@absons.ae
3. Verify **PDF files** exist in `EHC_Data_Pdf\` folder

#### CSV Download Issues
1. Check **network connectivity**
2. Verify **WiFi system** credentials
3. Check **download folder** permissions

## 📊 Monitoring

### Log Files
- `EHC_Logs\scheduler_setup_*.log` → Installation logs
- `EHC_Logs\email_evening_*.log` → Email logs  
- `EHC_Logs\download_files_*.log` → Download logs
- `EHC_Logs\vbs_upload_*.log` → Upload logs
- `EHC_Logs\vbs_report_*.log` → Report logs

### Success Indicators
- ✅ **CSV files** appear in `EHC_Data\{date}\`
- ✅ **Excel files** appear in `EHC_Data_Merge\{date}\`
- ✅ **PDF files** appear in `EHC_Data_Pdf\{date}\`
- ✅ **Email confirmations** in logs
- ✅ **Task Scheduler** shows "Ready" status

## 🎉 365-Day Reliability

This system is designed for **unattended operation** with:

- 🔄 **Automatic error recovery**
- 📊 **Comprehensive logging**  
- 🚀 **Process restart capabilities**
- 🔐 **State preservation across reboots**
- ⚡ **Power management integration**
- 🔒 **Lock/unlock state independence**

**Once installed, the system runs autonomously for 365 days without manual intervention.**
