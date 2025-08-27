# 🌙 MoonFlower Automation System

> **Complete 24/7 Windows Service for WiFi Data Processing & VBS Automation**

[![Windows Service](https://img.shields.io/badge/Windows-Service-blue.svg)](https://docs.microsoft.com/en-us/windows/win32/services/services)
[![Python](https://img.shields.io/badge/Python-3.7+-green.svg)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 📋 Table of Contents

- [🎯 Overview](#-overview)
- [✨ Key Features](#-key-features)
- [🏗️ System Architecture](#️-system-architecture)
- [📅 Daily Workflow](#-daily-workflow)
- [🚀 Quick Start](#-quick-start)
- [📦 Installation](#-installation)
- [🔧 Configuration](#-configuration)
- [📊 Components](#-components)
- [🎵 Audio Detection](#-audio-detection)
- [🤖 VBS Automation](#-vbs-automation)
- [📈 Monitoring](#-monitoring)
- [🛠️ Troubleshooting](#️-troubleshooting)
- [📝 API Reference](#-api-reference)
- [🤝 Contributing](#-contributing)

## 🎯 Overview

MoonFlower Automation is a professional-grade Windows service that automates the complete WiFi data processing workflow. It seamlessly integrates CSV downloads, Excel processing, VBS application automation, and PDF report generation into a single, reliable 24/7 background service.

### 🎪 What It Does

1. **📥 Automated CSV Downloads** - Downloads WiFi client data at scheduled times
2. **📊 Excel Processing** - Merges CSV files into VBS-compatible Excel format  
3. **🤖 VBS Automation** - Automates data upload using advanced image recognition
4. **🔊 Audio Detection** - Smart popup detection using sound analysis
5. **📄 PDF Generation** - Creates daily reports automatically
6. **📧 Email Delivery** - Sends reports to stakeholders
7. **🛡️ Service Mode** - Runs 24/7 without user interaction

## ✨ Key Features

### 🌟 Core Capabilities
- ✅ **True Windows Service** - Runs in background, survives user logoff/reboot
- ✅ **Smart Scheduling** - Precise timing for CSV downloads and processing  
- ✅ **Audio-Driven Automation** - Detects VBS popups using sound analysis
- ✅ **Image Recognition** - Advanced UI automation with confidence-based clicking
- ✅ **Error Recovery** - Intelligent error handling and automatic retry
- ✅ **Session 0 Compatible** - Works when PC is locked or no user logged in
- ✅ **One-Time Setup** - Install once, runs forever without admin prompts

### 🔥 Advanced Features
- 🎯 **TAB+ENTER Strategy** - Reliable VBS import button clicking
- 🔊 **Enhanced Audio Detection** - RMS, Peak, and Transient analysis
- 📊 **Status Tracking** - Prevents duplicate execution with daily reset
- ⏱️ **Extended Wait Times** - 15-minute import, 5-hour upload tolerance
- 🔄 **Process Management** - Automatic cleanup of stuck processes
- 📝 **Comprehensive Logging** - Detailed logs with timestamp precision

## 🏗️ System Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                 MoonFlower Automation Service               │
│                     (Windows Service)                      │
└─────────────────────┬───────────────────────────────────────┘
                      │
┌─────────────────────┼───────────────────────────────────────┐
│                     ▼                                       │
│  ┌─────────────┐ ┌─────────────┐ ┌─────────────┐ ┌─────────│──┐
│  │ CSV Module  │ │Excel Module │ │ VBS Module  │ │Audio Det││  │
│  │             │ │             │ │             │ │         ││  │
│  │• WiFi Data  │ │• File Merge │ │• Phase 1-4  │ │• Popup  ││  │
│  │• Downloads  │ │• VBS Format │ │• Image Rec  │ │• Sound  ││  │
│  │• Scheduler  │ │• Validation │ │• Automation │ │• RMS    ││  │
│  └─────────────┘ └─────────────┘ └─────────────┘ └─────────││──┘
└───────────────────────────────────────────────────────────┘│
                                                              │
┌─────────────────────────────────────────────────────────────┘
│  Support Components
│  ├── Path Manager (Centralized paths)
│  ├── Log Manager (Multi-level logging)  
│  ├── Error Handler (Recovery & retry)
│  ├── File Manager (Directory structure)
│  └── Config Manager (Settings & preferences)
└─────────────────────────────────────────────────────────────
```

## 📅 Daily Workflow

### ⏰ Automated Schedule

| Time | Action | Duration | Description |
|------|--------|----------|-------------|
| **9:30 AM** | 📥 CSV Download | ~5 min | First WiFi data collection |
| **12:30 PM** | 📥 CSV Download | ~5 min | Second WiFi data collection |
| **12:35 PM** | 📊 Excel Merge | ~2 min | Combine CSV files into Excel |
| **12:37 PM** | 🤖 VBS Phase 1 | ~2 min | Login to VBS application |
| **12:39 PM** | 🤖 VBS Phase 2 | ~1 min | Navigate to upload form |
| **12:40 PM** | 🤖 VBS Phase 3 | **15min-5hr** | Import data + Upload process |
| **Variable** | 🤖 VBS Phase 4 | ~10 min | Generate PDF report |
| **Complete** | 📧 Email Delivery | ~1 min | Send reports to stakeholders |

### 🔄 Phase 3 Deep Dive (Critical Phase)

```
Phase 3: Data Upload Process
├── Import EHC Checkbox ✓
├── Three Dots File Browser ✓  
├── Address Bar Navigation ✓
├── Excel File Selection ✓
├── Sheet Selection (EHC_Data) ✓
├── TAB+ENTER Import Strategy ✓
├── 🔊 Import Completion Audio Detection (15 min max)
├── Update Button Click ✓
├── 🔊 Upload Completion Audio Detection (5 hr max)
└── VBS Application Closure ✓
```

## 🚀 Quick Start

### 📋 Prerequisites

- **Windows 10/11** (Required for service mode)
- **Python 3.7+** with required packages
- **Admin rights** (one-time installation only)
- **VBS Application** installed and configured
- **Stable internet connection**

### ⚡ 5-Minute Setup

1. **Download the repository**
   ```bash
   git clone https://github.com/fasinabsons/Automata3.git
   cd Automata3
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Windows Service** (one-time admin required)
   ```cmd
   Right-click moonflower_automation.bat → "Run as Administrator"
   ```

4. **Verify installation**
   ```cmd
   sc query MoonFlowerAutomation
   ```

5. **🎉 Done!** - Service runs 24/7 automatically

## 📦 Installation

### 🔧 Detailed Installation Steps

#### Step 1: System Requirements
```bash
# Verify Python installation
python --version  # Should be 3.7+

# Check Windows version
winver  # Should be Windows 10 build 1903+ or Windows 11
```

#### Step 2: Clone Repository
```bash
git clone https://github.com/fasinabsons/Automata3.git
cd Automata3
```

#### Step 3: Install Dependencies
```bash
# Install core packages
pip install -r requirements.txt

# Install optional audio packages for enhanced detection
pip install scipy matplotlib
```

#### Step 4: Configure Paths
Edit `config/paths_config.json`:
```json
{
  "project_root": "C:/Users/YourUser/Documents/Automata3",
  "base_folders": {
    "csv_data": "EHC_Data",
    "excel_merge": "EHC_Data_Merge", 
    "pdf_reports": "EHC_Data_Pdf",
    "logs": "EHC_Logs"
  }
}
```

#### Step 5: Install Windows Service
```cmd
# Run as Administrator (ONE TIME ONLY)
Right-click moonflower_automation.bat → "Run as Administrator"

# Service installs automatically and starts immediately
```

## 🔧 Configuration

### 📁 Directory Structure
```
Automata3/
├── 📂 config/                    # Configuration files
│   ├── cv_config.json           # Computer vision settings
│   ├── settings.json            # General settings  
│   └── paths_config.json        # Path management
├── 📂 wifi/                     # CSV download components
│   ├── csv_downloader_simple.py # Main downloader
│   ├── element_detector.py      # UI detection
│   └── error_recovery.py        # Error handling
├── 📂 excel/                    # Excel processing
│   └── excel_generator.py       # CSV to Excel converter
├── 📂 vbs/                      # VBS automation
│   ├── vbs_phase1_login.py      # Login automation
│   ├── vbs_phase2_navigation_fixed.py # Navigation
│   ├── vbs_phase3_upload_complete.py # Upload process
│   ├── vbs_phase4_report_fixed.py    # Report generation
│   └── vbs_audio_detector.py    # Audio detection
├── 📂 utils/                    # Utility components
│   ├── path_manager.py          # Path management
│   ├── log_manager.py           # Logging system
│   ├── error_handler.py         # Error management
│   └── file_manager.py          # File operations
├── 📂 Images/                   # UI automation images
│   ├── phase1/                  # Login images
│   ├── phase2/                  # Navigation images
│   ├── phase3/                  # Upload images
│   └── phase4/                  # Report images
├── moonflower_automation.bat    # Main service script
└── README.md                    # This file
```

### ⚙️ Key Configuration Files

#### `config/cv_config.json` - Computer Vision Settings
```json
{
  "confidence_threshold": 0.8,
  "retry_attempts": 3,
  "click_delays": {
    "standard": 0.5,
    "critical": 1.0
  }
}
```

#### `config/settings.json` - General Settings  
```json
{
  "csv_slots": ["09:30", "12:30"],
  "excel_merge_delay": 5,
  "max_upload_hours": 5,
  "audio_detection": true
}
```

## 📊 Components

### 🌐 WiFi CSV Downloader (`wifi/`)

**Purpose**: Automated WiFi client data collection from web interface

**Key Features**:
- 🎯 Precise element detection using computer vision
- 🔄 Robust error recovery with exponential backoff
- ⏱️ Scheduled execution (9:30 AM, 12:30 PM)  
- 📊 Data validation and duplicate prevention
- 🔍 Smart waiting for dynamic content loading

**Main Files**:
- `csv_downloader_simple.py` - Core download logic
- `element_detector.py` - UI element detection  
- `error_recovery.py` - Error handling and retry logic

### 📈 Excel Generator (`excel/`)

**Purpose**: Convert CSV files to VBS-compatible Excel format

**Key Features**:
- 🔗 Intelligent column mapping
- 📁 Automatic file discovery and validation
- 🗓️ Date-based naming convention  
- 🧹 Data cleaning and normalization
- ✅ Output format validation

**Column Mapping**:
```python
csv_to_excel_mapping = {
    'Hostname': 'Hostname',
    'IP Address': 'IP_Address', 
    'MAC Address': 'MAC_Address',
    'WLAN (SSID)': 'Package',
    'AP MAC': 'AP_MAC',
    'Data Rate (up)': 'Upload',
    'Data Rate (down)': 'Download'
}
```

### 🤖 VBS Automation (`vbs/`)

**Purpose**: Automated interaction with VBS application

#### Phase 1: Login (`vbs_phase1_login.py`)
- 🚀 Application launch and window management
- 🔐 Automated login sequence
- ⏳ Smart waiting for application readiness
- 🖼️ Image-based UI element detection

#### Phase 2: Navigation (`vbs_phase2_navigation_fixed.py`)  
- 🧭 Menu navigation using keyboard shortcuts
- 🎯 Precise form location and activation
- 📍 Position-aware element clicking
- ⌨️ Fallback keyboard navigation

#### Phase 3: Upload (`vbs_phase3_upload_complete.py`)
- 📂 Advanced file browser automation
- 🔊 Audio-driven completion detection
- ⏱️ Extended wait times (15 min import, 5 hr upload)
- 🎯 TAB+ENTER import strategy for reliability
- 📱 VBS state monitoring ("Not Responding" = normal)

#### Phase 4: Report (`vbs_phase4_report_fixed.py`)
- 📊 Report generation automation  
- 🗓️ Precise date entry using triad navigation
- 💾 PDF export with filename management
- 📁 Address bar optimization for file saving

## 🎵 Audio Detection

### 🔊 Enhanced Audio Detection System

**Purpose**: Detect VBS application popups using audio analysis

**Technology Stack**:
- **PyAudio** - Real-time audio capture
- **NumPy** - Signal processing and analysis  
- **SciPy** - Advanced audio feature extraction
- **FFT Analysis** - Frequency domain processing

### 🎛️ Detection Methods

#### 1. RMS (Root Mean Square) Analysis
```python
rms_level = np.sqrt(np.mean(audio_frame**2))
if rms_level > threshold:
    popup_detected = True
```

#### 2. Peak Detection
```python
peaks = signal.find_peaks(audio_frame, height=peak_threshold)
if len(peaks[0]) > minimum_peaks:
    popup_detected = True
```

#### 3. Transient Detection  
```python
energy_diff = np.diff(np.square(audio_frame))
if np.max(energy_diff) > transient_threshold:
    popup_detected = True
```

### 🎯 Sound Tracking System

The system tracks 4 specific sounds during VBS automation:

| Sound Event | Phase | Trigger | Purpose |
|-------------|-------|---------|---------|
| **Three Dots 1** | Phase 3 | File browser open | Confirms file dialog opened |
| **Three Dots 2** | Phase 3 | Sheet selection | Confirms sheet dialog opened |
| **Import Success** | Phase 3 | Import completion | Triggers import OK click |
| **Upload Success** | Phase 3 | Upload completion | Triggers workflow completion |

### 🔧 Audio Configuration

```python
audio_config = {
    "sample_rate": 44100,          # CD quality
    "channels": 1,                 # Mono recording
    "chunk_size": 1024,           # Buffer size
    "detection_window": 5.0,       # Analysis window
    "rms_threshold": 0.01,         # Volume threshold
    "confidence_threshold": 0.7     # Overall confidence
}
```

## 🤖 VBS Automation

### 🎯 Image Recognition System

**Technology**: OpenCV + PyAutoGUI integration

**Features**:
- 🔍 Multi-confidence level detection (0.9, 0.8, 0.7)
- 📐 Precise click offset calculation  
- 🔄 Aggressive clicking for critical buttons
- ✅ Post-click verification (button disappears)
- ⚡ Fallback keyboard shortcuts

### 🎮 Click Strategies

#### Standard Click
```python
location = pyautogui.locateOnScreen(image, confidence=0.9)
if location:
    click_x, click_y = pyautogui.center(location) 
    pyautogui.click(click_x, click_y)
```

#### Aggressive Click (Critical Buttons)
```python
# Multiple attempts with decreasing confidence
for confidence in [0.9, 0.8, 0.7]:
    location = pyautogui.locateOnScreen(image, confidence=confidence)
    if location:
        pyautogui.click(pyautogui.center(location))
        pyautogui.click(pyautogui.center(location))  # Double click
        
        # Verify button disappeared
        if not pyautogui.locateOnScreen(image, confidence=0.8):
            return True  # Success!
```

#### TAB+ENTER Strategy
```python
# Most reliable for import button
pyautogui.typewrite("EHC_Data")  # Sheet name
pyautogui.press('tab')           # Navigate to import button  
pyautogui.press('enter')         # Click import button
```

### 🖱️ Advanced UI Techniques

#### Address Bar Navigation
```python
def navigate_using_address_bar(path):
    pyautogui.hotkey('ctrl', 'l')    # Focus address bar
    pyautogui.typewrite(path)        # Type full path
    pyautogui.press('enter')         # Navigate directly
```

#### Triad Date Navigation (Phase 4)
```python
def enter_date_triad(day, month, year):
    pyautogui.typewrite(day)
    pyautogui.press('right')     # Move to month field
    pyautogui.typewrite(month)  
    pyautogui.press('right')     # Move to year field
    pyautogui.typewrite(year)
```

## 📈 Monitoring

### 📊 Service Status Monitoring

#### Check Service Status
```cmd
# Basic status check
sc query MoonFlowerAutomation

# Detailed service information  
sc qc MoonFlowerAutomation

# Service startup type and state
Get-Service MoonFlowerAutomation | Format-List
```

#### Expected Output
```
SERVICE_NAME: MoonFlowerAutomation
TYPE               : 10  WIN32_OWN_PROCESS
STATE              : 4  RUNNING
WIN32_EXIT_CODE    : 0  (0x0)
SERVICE_EXIT_CODE  : 0  (0x0)
CHECKPOINT         : 0x0
WAIT_HINT          : 0x0
```

### 📝 Log Analysis

#### Log File Locations
```
📂 EHC_Logs/
├── 📂 25jul/                    # Today's logs
│   ├── service_20250725.log     # Service execution log
│   ├── vbs_phase3_complete_*.log # Phase 3 detailed log
│   └── automation_20250725.log  # General automation log
├── service_status.log           # Service-wide status log
└── daily_status.txt            # Current day status
```

#### Key Log Patterns to Monitor

**✅ Success Patterns**:
```log
[2025-07-25 09:30:15] [SERVICE] CSV SLOT 1: Completed successfully
[2025-07-25 12:30:15] [SERVICE] CSV SLOT 2: Completed successfully  
[2025-07-25 12:35:30] [SERVICE] EXCEL MERGE: Completed successfully
[2025-07-25 14:45:22] [SERVICE] VBS WORKFLOW: All phases completed successfully!
```

**⚠️ Warning Patterns**:
```log
[2025-07-25 12:40:15] INFO: ⏱️ Import wait: 5.0 minutes elapsed (max 15 minutes)
[2025-07-25 13:15:22] INFO: ⏱️ Upload: 1h 30m elapsed (max 2.0h remaining)
[2025-07-25 14:00:10] WARNING: VBS window no longer exists, searching again...
```

**❌ Error Patterns**:
```log
[2025-07-25 09:30:45] [SERVICE] ERROR: CSV SLOT 1: Failed - will retry next cycle
[2025-07-25 12:45:30] ERROR: VBS PHASE 3: FAILED - Upload process failed
[2025-07-25 15:00:00] ERROR: Critical error occurred - service recovery mode
```

### 📊 Daily Status Tracking

#### Status File Format (`daily_status.txt`)
```
25jul
csv_0930=completed
csv_1230=completed
excel_merge=completed  
vbs_workflow=pending
```

#### Status Values
- **`pending`** - Task not yet started
- **`completed`** - Task finished successfully
- **`failed`** - Task encountered error (will retry)

### 🔔 Monitoring Scripts

#### PowerShell Status Checker
```powershell
# Check service and recent activity
$service = Get-Service MoonFlowerAutomation
$logPath = "C:\Users\Lenovo\Documents\Automate2\Automata2\service_status.log"

Write-Host "Service Status: $($service.Status)"
Write-Host "Recent Activity:"
Get-Content $logPath -Tail 10
```

#### Python Status Monitor
```python
import subprocess
from pathlib import Path

def check_automation_status():
    # Check service status
    result = subprocess.run(['sc', 'query', 'MoonFlowerAutomation'], 
                          capture_output=True, text=True)
    
    service_running = "RUNNING" in result.stdout
    
    # Check daily status
    status_file = Path("daily_status.txt")
    if status_file.exists():
        with open(status_file) as f:
            status = dict(line.split('=') for line in f if '=' in line)
    
    return {
        "service_running": service_running,
        "daily_status": status,
        "all_complete": all(v == "completed" for v in status.values())
    }
```

## 🛠️ Troubleshooting

### 🚨 Common Issues & Solutions

#### Issue 1: Service Won't Start
**Symptoms**: 
```cmd
sc query MoonFlowerAutomation
# Shows: STATE: 1 STOPPED
```

**Solutions**:
```cmd
# 1. Check service installation
sc qc MoonFlowerAutomation

# 2. Try manual start with verbose output  
net start MoonFlowerAutomation

# 3. Check Windows Event Viewer
eventvwr.msc
# Navigate to: Windows Logs > System
# Look for MoonFlowerAutomation events

# 4. Reinstall service
moonflower_automation.bat uninstall
# Then reinstall as admin
```

#### Issue 2: VBS Automation Fails
**Symptoms**:
```log
VBS PHASE 3: FAILED - Upload process failed
❌ Could not locate: 07_import_button.png
```

**Solutions**:
```python
# 1. Verify image files exist
import os
image_path = "Images/phase3/07_import_button.png"
print(f"Image exists: {os.path.exists(image_path)}")

# 2. Test image detection manually
import pyautogui
location = pyautogui.locateOnScreen(image_path, confidence=0.7)
print(f"Image found at: {location}")

# 3. Check VBS window state
# Ensure VBS application is visible and not minimized

# 4. Update image files if UI changed
# Capture new screenshots using included tools
```

#### Issue 3: Audio Detection Not Working
**Symptoms**:
```log
WARN: Could not initialize audio detector
Audio detection error: No audio input device found
```

**Solutions**:
```cmd
# 1. Check audio devices
python -c "import pyaudio; pa=pyaudio.PyAudio(); [print(f'{i}: {pa.get_device_info_by_index(i)[\"name\"]}') for i in range(pa.get_device_count())]"

# 2. Install audio dependencies
pip install pyaudio scipy

# 3. Test microphone access
# Ensure no other applications are using microphone

# 4. Disable audio detection if not needed
# Edit vbs_phase3_upload_complete.py
# Set AUDIO_DETECTION_AVAILABLE = False
```

#### Issue 4: CSV Downloads Fail
**Symptoms**:
```log
CSV SLOT 1: Failed - will retry next cycle
ElementNotFound: Could not locate download button
```

**Solutions**:
```python
# 1. Check network connectivity
import requests
response = requests.get("https://google.com")
print(f"Network OK: {response.status_code == 200}")

# 2. Verify WiFi controller access
# Ensure VPN/firewall not blocking access

# 3. Update element detection images
# Web UI may have changed - capture new screenshots

# 4. Test manual download
# Try downloading CSV manually to verify process
```

### 🔍 Debug Mode

#### Enable Detailed Logging
```python
# Add to any script for enhanced debugging
import logging
logging.basicConfig(level=logging.DEBUG)

# Or modify log level in specific components
logger.setLevel(logging.DEBUG)
```

#### Capture Debug Information
```python
# Take screenshot for analysis
import pyautogui
screenshot = pyautogui.screenshot()
screenshot.save(f"debug_screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")

# Save window information
import win32gui
windows = []
win32gui.EnumWindows(lambda hwnd, results: results.append((hwnd, win32gui.GetWindowText(hwnd))), windows)
print("Available windows:", windows)
```

### 📞 Getting Help

#### Diagnostic Information to Collect
```cmd
# System information
systeminfo | findstr /B /C:"OS Name" /C:"OS Version" /C:"System Type"

# Service status
sc query MoonFlowerAutomation
sc qc MoonFlowerAutomation

# Python environment  
python --version
pip list | findstr "pyauto\|opencv\|numpy\|scipy"

# Recent logs (last 50 lines)
Get-Content "EHC_Logs\*\service_*.log" -Tail 50

# Disk space
dir C:\ | findstr "bytes free"
```

#### Log Package Creation
```powershell
# Create diagnostic package
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$packagePath = "MoonFlower_Diagnostics_$timestamp.zip"

# Collect log files
Compress-Archive -Path "EHC_Logs\*", "daily_status.txt", "service_status.log" -DestinationPath $packagePath

Write-Host "Diagnostic package created: $packagePath"
```

## 📝 API Reference

### 🔧 Core Components

#### PathManager Class
```python
from utils.path_manager import PathManager

pm = PathManager()
csv_dir = pm.get_csv_directory()      # Get today's CSV directory
excel_dir = pm.get_excel_directory()  # Get today's Excel directory
pdf_dir = pm.get_pdf_directory()      # Get today's PDF directory
log_dir = pm.get_logs_directory()     # Get today's log directory
```

#### LogManager Class
```python
from utils.log_manager import LogManager

lm = LogManager("component_name")
lm.info("Information message")
lm.warning("Warning message")  
lm.error("Error message")
lm.debug("Debug message")
```

#### Enhanced Audio Detector
```python
from vbs.vbs_audio_detector import EnhancedVBSAudioDetector

detector = EnhancedVBSAudioDetector(vbs_window_handle)
detector.initialize_audio_system()
detector.start_detection(timeout=300)  # 5 minutes

# Check for success sound
if detector.wait_for_success_sound(timeout=10):
    print("Popup detected!")
    
detector.cleanup()
```

### 🎯 VBS Automation Classes

#### VBS Phase 3 Complete
```python
from vbs.vbs_phase3_upload_complete import VBSPhase3Complete

phase3 = VBSPhase3Complete()
result = phase3.execute_complete_phase3()

if result["success"]:
    print(f"Upload completed in {result['execution_time_minutes']:.1f} minutes")
    print(f"Steps completed: {result['total_steps']}")
else:
    print(f"Upload failed: {result['error']}")
```

#### CSV Downloader
```python
from wifi.csv_downloader_simple import CSVDownloaderSimple

downloader = CSVDownloaderSimple()
result = downloader.download_csv_files()

if result["success"]:
    print(f"Downloaded {result['files_count']} CSV files")
    print(f"Saved to: {result['output_directory']}")
```

#### Excel Generator
```python
from excel.excel_generator import ExcelGenerator

generator = ExcelGenerator()
result = generator.merge_csv_to_excel()

if result["success"]:
    print(f"Excel file created: {result['output_file']}")
    print(f"Rows processed: {result['total_rows']}")
```

### 🛠️ Service Management API

#### Service Control
```python
import subprocess

def start_service():
    return subprocess.run(['net', 'start', 'MoonFlowerAutomation'])

def stop_service():
    return subprocess.run(['net', 'stop', 'MoonFlowerAutomation'])

def service_status():
    result = subprocess.run(['sc', 'query', 'MoonFlowerAutomation'], 
                          capture_output=True, text=True)
    return "RUNNING" in result.stdout
```

#### Status Management
```python
from pathlib import Path

def get_daily_status():
    status_file = Path("daily_status.txt")
    if status_file.exists():
        with open(status_file) as f:
            lines = f.read().strip().split('\n')
            date_folder = lines[0]
            status = {}
            for line in lines[1:]:
                if '=' in line:
                    key, value = line.split('=', 1)
                    status[key] = value
            return {"date": date_folder, "status": status}
    return None

def set_task_status(task, status):
    """Set status for a specific task (csv_0930, csv_1230, excel_merge, vbs_workflow)"""
    current = get_daily_status()
    if current:
        current["status"][task] = status
        with open("daily_status.txt", "w") as f:
            f.write(f"{current['date']}\n")
            for key, value in current["status"].items():
                f.write(f"{key}={value}\n")
```

## 🤝 Contributing

### 🎯 How to Contribute

1. **Fork the repository**
   ```bash
   git clone https://github.com/fasinabsons/Automata3.git
   cd Automata3
   ```

2. **Create feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

3. **Make changes and test**
   ```bash
   # Test your changes thoroughly
   python -m pytest tests/
   ```

4. **Commit with clear message**
   ```bash
   git commit -m "feat: add new audio detection method"
   ```

5. **Push and create pull request**
   ```bash
   git push origin feature/your-feature-name
   ```

### 🧪 Development Setup

#### Local Development Environment
```bash
# Clone repository
git clone https://github.com/fasinabsons/Automata3.git
cd Automata3

# Create virtual environment
python -m venv moonflower_env
moonflower_env\Scripts\activate

# Install development dependencies
pip install -r requirements.txt
pip install -r requirements-dev.txt

# Install pre-commit hooks
pre-commit install
```

#### Testing Framework
```bash
# Run all tests
python -m pytest

# Run specific test category
python -m pytest tests/test_audio_detection.py
python -m pytest tests/test_vbs_automation.py
python -m pytest tests/test_csv_download.py

# Run with coverage
python -m pytest --cov=. --cov-report=html
```

### 📋 Contribution Guidelines

#### Code Style
- **PEP 8** compliance for Python code
- **Type hints** for function parameters and returns
- **Docstrings** for all public methods and classes
- **Error handling** with specific exception types

#### Commit Message Format
```
type(scope): brief description

Longer description if needed

Fixes #issue_number
```

**Types**: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

#### Pull Request Checklist
- [ ] Code follows project style guidelines
- [ ] Tests added for new functionality  
- [ ] Documentation updated
- [ ] No breaking changes (or clearly documented)
- [ ] Service installation tested
- [ ] VBS automation verified

### 🏗️ Architecture Principles

#### Design Philosophy
1. **Reliability First** - Service must run 24/7 without intervention
2. **Error Recovery** - Graceful handling of all failure scenarios
3. **Modular Design** - Each component independently testable
4. **Extensive Logging** - Full audit trail of all operations
5. **Smart Automation** - Intelligent detection and decision making

#### Code Organization
```
├── 📂 Core Components/
│   ├── Service Layer (Windows service integration)
│   ├── Automation Layer (VBS, CSV, Excel)  
│   ├── Detection Layer (Audio, Image, UI)
│   └── Infrastructure Layer (Logging, Paths, Config)
├── 📂 External Interfaces/
│   ├── VBS Application Integration
│   ├── WiFi Controller Web Interface
│   └── File System Operations
└── 📂 Support Systems/
│   ├── Error Recovery & Retry Logic
│   ├── Audio Processing Pipeline
│   └── Service Management & Monitoring
```

---

## 📜 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **PyAutoGUI** - UI automation framework
- **OpenCV** - Computer vision and image processing
- **PyAudio** - Real-time audio processing  
- **NumPy & SciPy** - Numerical computing and signal analysis
- **NSSM** - Non-Sucking Service Manager for Windows services

---

## 📞 Support & Contact

- **GitHub Issues**: [Create an issue](https://github.com/fasinabsons/Automata3/issues)
- **Documentation**: [Wiki](https://github.com/fasinabsons/Automata3/wiki)
- **Discussions**: [GitHub Discussions](https://github.com/fasinabsons/Automata3/discussions)

---

<div align="center">

### 🌙 MoonFlower Automation
**Professional WiFi Data Processing & VBS Automation**

[![GitHub stars](https://img.shields.io/github/stars/fasinabsons/Automata3.svg)](https://github.com/fasinabsons/Automata3/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/fasinabsons/Automata3.svg)](https://github.com/fasinabsons/Automata3/network)
[![GitHub issues](https://img.shields.io/github/issues/fasinabsons/Automata3.svg)](https://github.com/fasinabsons/Automata3/issues)

Made with ❤️ for automated WiFi data processing

</div>
