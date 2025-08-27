# MoonFlower Automation - Windows Service Installation Guide

## 🚀 One-Time Installation (Requires Admin Rights)

### Step 1: Install the Service
1. **Right-click** on `moonflower_automation.bat`
2. **Select "Run as Administrator"** (this is the ONLY time you need admin rights)
3. The system will automatically:
   - Download service wrapper (NSSM)
   - Install Windows service
   - Configure background execution
   - Start the service immediately

### Step 2: Verify Installation
After installation, you'll see:
```
✅ Service Installation Complete!

The MoonFlower Automation service is now installed and running.
It will automatically start with Windows and run in the background,
even when the PC is locked or no user is logged in.

Service Features:
• Runs 24/7 without user interaction
• No admin prompts after installation
• Works when PC is locked
• Automatic startup with Windows
• Background Session 0 execution
```

## 🔄 Service Management

### Start the Service
```cmd
net start MoonFlowerAutomation
```

### Stop the Service
```cmd
net stop MoonFlowerAutomation
```

### Check Service Status
```cmd
sc query MoonFlowerAutomation
```

### Uninstall the Service
```cmd
moonflower_automation.bat uninstall
```

## 🎯 How It Works

### Daily Schedule
- **9:30 AM** → CSV Download (Slot 1)
- **12:30 PM** → CSV Download (Slot 2)  
- **12:35 PM** → Excel Merge
- **After Excel** → VBS Workflow (Login → Navigation → Upload → PDF)

### Background Operation
- ✅ **Runs when PC is locked**
- ✅ **No user interaction required**
- ✅ **Automatic startup with Windows**
- ✅ **Session 0 isolation (background service)**
- ✅ **No admin prompts after installation**

### Logging
- **Service logs**: `EHC_Logs/[date]/service_YYYYMMDD.log`
- **Status tracking**: `daily_status.txt`
- **Service status**: `service_status.log`

## 🛡️ Service Features

### Security & Reliability
- Runs as Windows service (most reliable)
- Automatic process cleanup on errors
- 7-day restart cycle for maintenance
- Protected from user logoff/lockscreen

### Smart Execution
- Only executes when scheduled times arrive
- Prevents duplicate execution with status tracking
- Automatically resets daily at midnight
- Continues from last successful step on restart

### Error Recovery
- 5-minute recovery wait on errors
- Automatic process termination cleanup
- Service restart capability
- Comprehensive error logging

## 📊 Monitoring

### Check if Service is Running
```cmd
sc query MoonFlowerAutomation
```
Should show: `STATE: 4 RUNNING`

### View Recent Logs
Check the latest log file in:
```
C:\Users\Lenovo\Documents\Automate2\Automata2\EHC_Logs\[today's date]\
```

### Check Daily Status
View `daily_status.txt` to see completion status:
```
25jul
csv_0930=completed
csv_1230=completed  
excel_merge=completed
vbs_workflow=pending
```

## 🔧 Troubleshooting

### Service Won't Start
1. Check if service is installed: `sc query MoonFlowerAutomation`
2. Try manual start: `net start MoonFlowerAutomation`
3. Check Windows Event Viewer for service errors

### VBS Automation Not Working
1. Ensure VBS application is installed and accessible
2. Check that Python scripts exist in the correct paths
3. Verify image files are present in `Images/phase3/` etc.

### No CSV Downloads
1. Check network connectivity
2. Verify WiFi automation scripts are working
3. Check logs for download errors

## 🎉 Benefits of Service Mode

### vs Regular BAT File
- ❌ **Regular**: Stops when user logs off
- ✅ **Service**: Continues running always

### vs Task Scheduler  
- ❌ **Task Scheduler**: Limited scheduling options
- ✅ **Service**: Continuous monitoring with smart timing

### vs Startup Programs
- ❌ **Startup**: Requires user login
- ✅ **Service**: Runs without any user interaction

## 📞 Support

If you encounter issues:
1. Check the service logs in `EHC_Logs/[date]/`
2. Verify service status with `sc query MoonFlowerAutomation`
3. Try stopping and restarting the service
4. If needed, uninstall and reinstall the service

---

**Remember**: After the one-time installation, the system runs completely automatically without any admin prompts or user interaction required! 