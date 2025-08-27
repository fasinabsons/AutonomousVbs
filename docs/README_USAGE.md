# MoonFlower Automated CSV Download System v3.0

## 🚀 Quick Start (Recommended)

**Just run the BAT file as Administrator - that's it!**

1. Right-click `csv_download.bat`
2. Select "Run as administrator"
3. The system will automatically:
   - Download CSV files at scheduled times (8:30, 9:30, 10:30, 11:30, 12:30, 13:30, 14:30, 15:30, 16:30)
   - Merge CSV files into Excel after downloads complete
   - Create PDF directory structure for VBS application
   - Send email notification to faseenm@gmail.com when done

## ⏰ Automatic Schedule

### Morning Session (8:00 AM - 12:59 PM)
- **Download Times**: 8:30, 9:30, 10:30, 11:30, **12:30**
- **Files**: Saved to `EHC_Data/[date]/`

### Afternoon Session (1:00 PM - 5:00 PM)  
- **Download Times**: 13:30, 14:30, 15:30, 16:30
- **Files**: Saved to `EHC_Data/[date]/`

### After All Downloads Complete
- **Excel Merge**: Creates consolidated Excel file in `EHC_Data_Merge/[date]/`
- **PDF Directory**: Creates `EHC_Data_Pdf/[date]/` for VBS application
- **Email**: Sends completion notification to faseenm@gmail.com

## 📁 Directory Structure Created

```
Automata2/
├── EHC_Data/22jul/           ← CSV files downloaded here
├── EHC_Data_Merge/22jul/     ← Excel files created here  
├── EHC_Data_Pdf/22jul/       ← PDF directory for VBS app
└── EHC_Logs/22jul/           ← Log files for troubleshooting
```

## 🎯 What Happens at 12:30 PM

When you run the BAT file as admin:

1. **10:45 AM** - System starts, waits for next scheduled time
2. **12:30 PM** - Automatically downloads CSV files from all network slots
3. **12:35 PM** - Download completes, waits for afternoon session
4. **4:30 PM** - Downloads afternoon CSV files  
5. **4:35 PM** - All downloads complete, starts Excel merge
6. **4:36 PM** - Creates PDF directory, sends email to faseenm@gmail.com
7. **4:37 PM** - ✅ **ALL DONE!** System continues monitoring for next day

## 🔧 Manual Options (if needed)

If you need to run manually instead of automatic:

```batch
csv_download.bat /now          # Download immediately and merge
csv_download.bat /morning      # Morning session only
csv_download.bat /afternoon    # Afternoon session only  
csv_download.bat /help         # Show all options
```

## 📧 Email Notification

You'll receive an email at **faseenm@gmail.com** with:
- ✅ Download summary (files downloaded, success rate)
- 📊 Excel file location
- 📁 Directory paths
- ⚠️ Any warnings or errors
- 📄 PDF directory confirmation

## 🛠️ Troubleshooting

### If downloads fail:
1. Check `EHC_Logs/[date]/` for error details
2. Ensure Chrome browser is installed
3. Check internet connection
4. Run `csv_download.bat /now` for immediate retry

### If email doesn't arrive:
- Check spam folder
- Email is sent to: faseenm@gmail.com
- Check logs for email delivery status

## ✨ Key Features

- **🔄 Fully Automated** - Just run once, handles everything
- **⏰ Time-Based** - Downloads at optimal slot times
- **🔧 Self-Healing** - Automatically installs missing dependencies
- **📊 Excel Integration** - Auto-merges CSV files
- **📧 Email Notifications** - Keeps you informed
- **📁 VBS Ready** - Creates PDF directory structure
- **🛡️ Error Recovery** - Retries failed operations
- **📝 Comprehensive Logging** - Full audit trail

## 🎉 Success Indicators

When everything works correctly, you'll see:
- ✅ CSV files in `EHC_Data/22jul/`
- ✅ Excel file in `EHC_Data_Merge/22jul/`  
- ✅ PDF directory `EHC_Data_Pdf/22jul/` created
- ✅ Email received at faseenm@gmail.com
- ✅ Log files show "SUCCESS" messages

---

**That's it! The system is designed to be completely hands-off once you start it as admin.**