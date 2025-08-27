# VBS OpenCV Modernization - Deployment Guide

## Overview

This guide provides step-by-step instructions for deploying the VBS OpenCV modernization system in production environments. It covers installation, configuration, testing, and maintenance procedures.

## Prerequisites

### System Requirements

#### Minimum Requirements
- **Operating System**: Windows 10 or Windows Server 2016+
- **RAM**: 8 GB minimum, 16 GB recommended
- **CPU**: Quad-core processor, 2.5 GHz or higher
- **Storage**: 10 GB free space for installation and logs
- **Display**: 1920x1080 resolution minimum
- **Network**: Internet connection for initial setup

#### Recommended Requirements
- **Operating System**: Windows 11 or Windows Server 2022
- **RAM**: 32 GB for optimal performance
- **CPU**: 8-core processor, 3.0 GHz or higher
- **Storage**: SSD with 50 GB free space
- **Display**: 2560x1440 resolution or higher
- **Network**: High-speed internet connection

### Software Dependencies

#### Required Software
1. **Python 3.11+**
   - Download from: https://www.python.org/downloads/
   - Ensure "Add Python to PATH" is checked during installation

2. **Tesseract OCR**
   - Download from: https://github.com/UB-Mannheim/tesseract/wiki
   - Install to default location: `C:\Program Files\Tesseract-OCR\`

3. **Microsoft Visual C++ Redistributable**
   - Download latest version from Microsoft
   - Required for OpenCV and other native libraries

4. **VBS Application**
   - Ensure VBS application is installed and accessible
   - Verify application launches and functions normally

## Installation Process

### Step 1: Environment Setup

#### 1.1 Create Installation Directory
```cmd
mkdir C:\VBS_OpenCV_Automation
cd C:\VBS_OpenCV_Automation
```

#### 1.2 Download and Extract Files
```cmd
# Extract the VBS OpenCV modernization files to the installation directory
# Ensure the following structure exists:
# C:\VBS_OpenCV_Automation\
# ├── vbs/
# ├── cv_services/
# ├── config/
# ├── docs/
# ├── requirements.txt
# └── README.md
```

#### 1.3 Create Python Virtual Environment
```cmd
python -m venv vbs_opencv_env
vbs_opencv_env\Scripts\activate
```

### Step 2: Install Python Dependencies

#### 2.1 Upgrade pip
```cmd
python -m pip install --upgrade pip
```

#### 2.2 Install Required Packages
```cmd
pip install -r requirements.txt
```

#### 2.3 Verify Installation
```cmd
python -c "import cv2, pytesseract, numpy, pillow; print('All packages installed successfully')"
```

### Step 3: Configure Tesseract OCR

#### 3.1 Verify Tesseract Installation
```cmd
tesseract --version
```

#### 3.2 Test Tesseract Functionality
```cmd
# Create a test image with text
echo "Test OCR functionality" > test.txt
# Use Tesseract to read the text (manual verification)
```

### Step 4: Configure the System

#### 4.1 Update Configuration File
Edit `config/cv_config.json`:

```json
{
  "ocr_settings": {
    "tesseract_path": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
    "language": "eng",
    "confidence_threshold": 0.7,
    "page_segmentation_mode": 6,
    "character_whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .-_()[]{}:;,!?@#$%&*+=<>/\\|\"'`~",
    "preprocessing": {
      "gaussian_blur_kernel": 3,
      "adaptive_threshold_block_size": 11,
      "adaptive_threshold_c": 2,
      "morphology_kernel_size": 2
    }
  },
  "template_matching": {
    "confidence_threshold": 0.8,
    "template_cache_size": 50,
    "max_template_variations": 5,
    "scale_factors": [0.8, 0.9, 1.0, 1.1, 1.2],
    "template_directory": "vbs/templates"
  },
  "smart_automation": {
    "method_priority": ["ocr", "template", "coordinates"],
    "max_retries": 3,
    "retry_delay": 1.0,
    "exponential_backoff": true,
    "screenshot_on_error": true,
    "performance_tracking": true
  },
  "vbs_specific": {
    "window_title_patterns": ["absons", "arabian", "moonflower", "erp"],
    "exclude_window_patterns": ["login", "security", "warning", "browser"],
    "import_timeout_minutes": 5,
    "update_timeout_minutes": 30,
    "popup_detection_timeout": 10,
    "ui_response_delay": 0.5
  },
  "performance": {
    "screenshot_region_optimization": true,
    "parallel_processing": false,
    "cache_successful_locations": true,
    "cache_duration_seconds": 300,
    "max_concurrent_operations": 2
  },
  "debugging": {
    "debug_mode": false,
    "save_debug_images": false,
    "debug_image_path": "debug_images",
    "verbose_logging": false,
    "screenshot_failed_operations": true
  }
}
```

#### 4.2 Create Required Directories
```cmd
mkdir EHC_Logs
mkdir debug_images
mkdir vbs\templates
mkdir vbs\templates\navigation
mkdir vbs\templates\upload
mkdir vbs\templates\reports
mkdir vbs\templates\common
```

#### 4.3 Set Up Template Images
Copy template images to appropriate directories:
- Navigation templates → `vbs/templates/navigation/`
- Upload templates → `vbs/templates/upload/`
- Report templates → `vbs/templates/reports/`
- Common templates → `vbs/templates/common/`

### Step 5: Initial Testing

#### 5.1 Test Computer Vision Services
```cmd
python -c "
from cv_services.ocr_service import OCRService
from cv_services.template_service import TemplateService
from cv_services.smart_engine import SmartAutomationEngine

print('Testing OCR Service...')
ocr = OCRService()
print('OCR Service: OK')

print('Testing Template Service...')
template = TemplateService()
print('Template Service: OK')

print('Testing Smart Engine...')
engine = SmartAutomationEngine()
print('Smart Engine: OK')

print('All services initialized successfully!')
"
```

#### 5.2 Test VBS Integration
```cmd
python -c "
from vbs.vbs_phase2_navigation import VBSPhase2_Navigation

print('Testing VBS Phase 2 Navigation...')
nav = VBSPhase2_Navigation()
print('VBS Phase 2: OK')

print('VBS integration test completed successfully!')
"
```

## Production Configuration

### Security Configuration

#### 5.1 File Permissions
```cmd
# Set appropriate permissions for the installation directory
icacls C:\VBS_OpenCV_Automation /grant:r "Users:(OI)(CI)RX"
icacls C:\VBS_OpenCV_Automation\config /grant:r "Users:(OI)(CI)RX"
icacls C:\VBS_OpenCV_Automation\EHC_Logs /grant:r "Users:(OI)(CI)F"
```

#### 5.2 Windows Firewall
```cmd
# Allow Python through Windows Firewall if needed
netsh advfirewall firewall add rule name="VBS OpenCV Automation" dir=in action=allow program="C:\VBS_OpenCV_Automation\vbs_opencv_env\Scripts\python.exe"
```

### Performance Configuration

#### 5.1 Production Configuration Settings
Update `config/cv_config.json` for production:

```json
{
  "debugging": {
    "debug_mode": false,
    "save_debug_images": false,
    "verbose_logging": false,
    "screenshot_failed_operations": true
  },
  "performance": {
    "screenshot_region_optimization": true,
    "parallel_processing": true,
    "cache_successful_locations": true,
    "cache_duration_seconds": 600,
    "max_concurrent_operations": 3
  }
}
```

#### 5.2 Windows Performance Settings
```cmd
# Set high performance power plan
powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c

# Disable Windows visual effects for better performance
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" /v VisualFXSetting /t REG_DWORD /d 2 /f
```

## Service Installation

### Windows Service Setup

#### 6.1 Create Service Wrapper
Create `vbs_opencv_service.py`:

```python
#!/usr/bin/env python3
"""
VBS OpenCV Automation Windows Service
"""

import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import os
import time
import logging
from pathlib import Path

class VBSOpenCVService(win32serviceutil.ServiceFramework):
    _svc_name_ = "VBSOpenCVAutomation"
    _svc_display_name_ = "VBS OpenCV Automation Service"
    _svc_description_ = "Automated VBS operations using computer vision"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        
        # Setup logging
        self.setup_logging()
        
    def setup_logging(self):
        """Setup service logging"""
        log_path = Path("EHC_Logs/vbs_opencv_service.log")
        log_path.parent.mkdir(exist_ok=True)
        
        logging.basicConfig(
            filename=str(log_path),
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        self.logger = logging.getLogger(__name__)
    
    def SvcStop(self):
        """Stop the service"""
        self.logger.info("VBS OpenCV Service stopping...")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
    
    def SvcDoRun(self):
        """Run the service"""
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        
        self.logger.info("VBS OpenCV Service starting...")
        self.main()
    
    def main(self):
        """Main service loop"""
        try:
            # Import VBS automation modules
            from orchestrator import main as run_orchestrator
            
            self.logger.info("Starting VBS automation orchestrator...")
            
            # Run the main automation loop
            while True:
                # Check if service should stop
                if win32event.WaitForSingleObject(self.hWaitStop, 1000) == win32event.WAIT_OBJECT_0:
                    self.logger.info("Service stop requested")
                    break
                
                try:
                    # Run automation cycle
                    run_orchestrator()
                    
                except Exception as e:
                    self.logger.error(f"Automation cycle error: {e}")
                    time.sleep(60)  # Wait before retry
                
                # Wait between cycles
                time.sleep(300)  # 5 minutes
                
        except Exception as e:
            self.logger.error(f"Service error: {e}")
            servicemanager.LogErrorMsg(f"VBS OpenCV Service error: {e}")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(VBSOpenCVService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(VBSOpenCVService)
```

#### 6.2 Install Windows Service
```cmd
# Install the service
python vbs_opencv_service.py install

# Start the service
python vbs_opencv_service.py start

# Check service status
python vbs_opencv_service.py status
```

### Task Scheduler Alternative

If Windows Service is not preferred, use Task Scheduler:

#### 6.3 Create Scheduled Task
```cmd
# Create a batch file to run the automation
echo @echo off > run_vbs_automation.bat
echo cd /d "C:\VBS_OpenCV_Automation" >> run_vbs_automation.bat
echo vbs_opencv_env\Scripts\activate >> run_vbs_automation.bat
echo python orchestrator.py >> run_vbs_automation.bat

# Create scheduled task
schtasks /create /tn "VBS OpenCV Automation" /tr "C:\VBS_OpenCV_Automation\run_vbs_automation.bat" /sc daily /st 08:00 /ru SYSTEM
```

## Monitoring and Maintenance

### Log Monitoring

#### 7.1 Log File Locations
- Service logs: `EHC_Logs/vbs_opencv_service.log`
- OCR logs: `EHC_Logs/ocr_service.log`
- Smart engine logs: `EHC_Logs/smart_engine.log`
- Error handler logs: `EHC_Logs/cv_error_handler.log`
- VBS phase logs: `EHC_Logs/vbs_phase*.log`

#### 7.2 Log Rotation Setup
Create `setup_log_rotation.py`:

```python
import os
import glob
import time
from datetime import datetime, timedelta

def rotate_logs(log_directory="EHC_Logs", max_age_days=30, max_size_mb=50):
    """Rotate log files based on age and size"""
    
    log_files = glob.glob(os.path.join(log_directory, "*.log"))
    
    for log_file in log_files:
        try:
            # Check file age
            file_time = os.path.getmtime(log_file)
            file_age = datetime.now() - datetime.fromtimestamp(file_time)
            
            # Check file size
            file_size_mb = os.path.getsize(log_file) / (1024 * 1024)
            
            # Rotate if too old or too large
            if file_age.days > max_age_days or file_size_mb > max_size_mb:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_name = f"{log_file}.{timestamp}"
                os.rename(log_file, backup_name)
                print(f"Rotated {log_file} to {backup_name}")
        
        except Exception as e:
            print(f"Error rotating {log_file}: {e}")

if __name__ == "__main__":
    rotate_logs()
```

### Performance Monitoring

#### 7.3 Create Performance Monitor Script
Create `monitor_performance.py`:

```python
#!/usr/bin/env python3
"""
VBS OpenCV Performance Monitor
"""

import time
import psutil
import json
import logging
from datetime import datetime
from cv_services.smart_engine import SmartAutomationEngine

class PerformanceMonitor:
    def __init__(self):
        self.setup_logging()
        self.smart_engine = SmartAutomationEngine()
    
    def setup_logging(self):
        logging.basicConfig(
            filename="EHC_Logs/performance_monitor.log",
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger(__name__)
    
    def collect_system_metrics(self):
        """Collect system performance metrics"""
        return {
            "timestamp": datetime.now().isoformat(),
            "cpu_percent": psutil.cpu_percent(interval=1),
            "memory_percent": psutil.virtual_memory().percent,
            "disk_usage": psutil.disk_usage('C:').percent,
            "process_count": len(psutil.pids())
        }
    
    def collect_automation_metrics(self):
        """Collect automation performance metrics"""
        try:
            stats = self.smart_engine.get_performance_stats()
            return {
                "timestamp": datetime.now().isoformat(),
                "total_operations": stats.get("total_operations", 0),
                "success_rate": stats.get("overall_success_rate", 0),
                "avg_execution_time": stats.get("average_execution_times", {}).get("ocr", 0),
                "cache_hit_rate": stats.get("cache_hit_rate", 0)
            }
        except Exception as e:
            self.logger.error(f"Error collecting automation metrics: {e}")
            return {}
    
    def check_alerts(self, system_metrics, automation_metrics):
        """Check for performance alerts"""
        alerts = []
        
        # System alerts
        if system_metrics["cpu_percent"] > 80:
            alerts.append(f"High CPU usage: {system_metrics['cpu_percent']:.1f}%")
        
        if system_metrics["memory_percent"] > 85:
            alerts.append(f"High memory usage: {system_metrics['memory_percent']:.1f}%")
        
        if system_metrics["disk_usage"] > 90:
            alerts.append(f"High disk usage: {system_metrics['disk_usage']:.1f}%")
        
        # Automation alerts
        if automation_metrics.get("success_rate", 1) < 0.8:
            alerts.append(f"Low automation success rate: {automation_metrics['success_rate']:.1%}")
        
        if automation_metrics.get("avg_execution_time", 0) > 10:
            alerts.append(f"Slow automation execution: {automation_metrics['avg_execution_time']:.1f}s")
        
        return alerts
    
    def run_monitoring_cycle(self):
        """Run one monitoring cycle"""
        try:
            # Collect metrics
            system_metrics = self.collect_system_metrics()
            automation_metrics = self.collect_automation_metrics()
            
            # Check for alerts
            alerts = self.check_alerts(system_metrics, automation_metrics)
            
            # Log metrics
            self.logger.info(f"System metrics: {json.dumps(system_metrics)}")
            self.logger.info(f"Automation metrics: {json.dumps(automation_metrics)}")
            
            # Log alerts
            for alert in alerts:
                self.logger.warning(f"ALERT: {alert}")
            
            return len(alerts) == 0
            
        except Exception as e:
            self.logger.error(f"Monitoring cycle error: {e}")
            return False
    
    def run(self, interval_minutes=5):
        """Run continuous monitoring"""
        self.logger.info("Performance monitoring started")
        
        while True:
            try:
                self.run_monitoring_cycle()
                time.sleep(interval_minutes * 60)
            except KeyboardInterrupt:
                self.logger.info("Performance monitoring stopped")
                break
            except Exception as e:
                self.logger.error(f"Monitoring error: {e}")
                time.sleep(60)

if __name__ == "__main__":
    monitor = PerformanceMonitor()
    monitor.run()
```

## Backup and Recovery

### Backup Strategy

#### 8.1 Create Backup Script
Create `backup_system.py`:

```python
#!/usr/bin/env python3
"""
VBS OpenCV System Backup
"""

import os
import shutil
import zipfile
import json
from datetime import datetime
from pathlib import Path

class SystemBackup:
    def __init__(self, backup_dir="backups"):
        self.backup_dir = Path(backup_dir)
        self.backup_dir.mkdir(exist_ok=True)
    
    def create_backup(self):
        """Create complete system backup"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"vbs_opencv_backup_{timestamp}"
        backup_path = self.backup_dir / f"{backup_name}.zip"
        
        with zipfile.ZipFile(backup_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Backup configuration
            self._backup_directory(zipf, "config", "config/")
            
            # Backup templates
            self._backup_directory(zipf, "vbs/templates", "templates/")
            
            # Backup logs (last 7 days)
            self._backup_recent_logs(zipf, "EHC_Logs", days=7)
            
            # Backup custom scripts
            for script in ["orchestrator.py", "vbs_opencv_service.py"]:
                if os.path.exists(script):
                    zipf.write(script, f"scripts/{script}")
        
        print(f"Backup created: {backup_path}")
        return backup_path
    
    def _backup_directory(self, zipf, source_dir, archive_prefix):
        """Backup entire directory"""
        if os.path.exists(source_dir):
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    archive_path = os.path.join(archive_prefix, os.path.relpath(file_path, source_dir))
                    zipf.write(file_path, archive_path)
    
    def _backup_recent_logs(self, zipf, log_dir, days=7):
        """Backup recent log files"""
        if not os.path.exists(log_dir):
            return
        
        cutoff_time = datetime.now().timestamp() - (days * 24 * 3600)
        
        for file in os.listdir(log_dir):
            file_path = os.path.join(log_dir, file)
            if os.path.isfile(file_path) and os.path.getmtime(file_path) > cutoff_time:
                zipf.write(file_path, f"logs/{file}")
    
    def cleanup_old_backups(self, keep_days=30):
        """Remove old backup files"""
        cutoff_time = datetime.now().timestamp() - (keep_days * 24 * 3600)
        
        for backup_file in self.backup_dir.glob("vbs_opencv_backup_*.zip"):
            if backup_file.stat().st_mtime < cutoff_time:
                backup_file.unlink()
                print(f"Removed old backup: {backup_file}")

if __name__ == "__main__":
    backup = SystemBackup()
    backup.create_backup()
    backup.cleanup_old_backups()
```

#### 8.2 Schedule Automated Backups
```cmd
# Create scheduled task for daily backups
schtasks /create /tn "VBS OpenCV Backup" /tr "C:\VBS_OpenCV_Automation\vbs_opencv_env\Scripts\python.exe C:\VBS_OpenCV_Automation\backup_system.py" /sc daily /st 02:00
```

### Recovery Procedures

#### 8.3 System Recovery Script
Create `restore_system.py`:

```python
#!/usr/bin/env python3
"""
VBS OpenCV System Recovery
"""

import os
import zipfile
import shutil
from pathlib import Path

class SystemRestore:
    def __init__(self):
        self.backup_dir = Path("backups")
    
    def list_backups(self):
        """List available backups"""
        backups = list(self.backup_dir.glob("vbs_opencv_backup_*.zip"))
        backups.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        
        print("Available backups:")
        for i, backup in enumerate(backups):
            timestamp = backup.stem.split("_")[-2:]
            date_str = f"{timestamp[0][:4]}-{timestamp[0][4:6]}-{timestamp[0][6:8]}"
            time_str = f"{timestamp[1][:2]}:{timestamp[1][2:4]}:{timestamp[1][4:6]}"
            print(f"{i+1}. {backup.name} ({date_str} {time_str})")
        
        return backups
    
    def restore_from_backup(self, backup_path):
        """Restore system from backup"""
        if not os.path.exists(backup_path):
            print(f"Backup file not found: {backup_path}")
            return False
        
        print(f"Restoring from backup: {backup_path}")
        
        try:
            with zipfile.ZipFile(backup_path, 'r') as zipf:
                # Restore configuration
                self._restore_directory(zipf, "config/", "config")
                
                # Restore templates
                self._restore_directory(zipf, "templates/", "vbs/templates")
                
                # Restore scripts
                self._restore_directory(zipf, "scripts/", ".")
            
            print("System restored successfully")
            return True
            
        except Exception as e:
            print(f"Restore failed: {e}")
            return False
    
    def _restore_directory(self, zipf, archive_prefix, target_dir):
        """Restore directory from archive"""
        os.makedirs(target_dir, exist_ok=True)
        
        for member in zipf.namelist():
            if member.startswith(archive_prefix):
                target_path = os.path.join(target_dir, member[len(archive_prefix):])
                
                # Create directory if needed
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                
                # Extract file
                with zipf.open(member) as source, open(target_path, 'wb') as target:
                    shutil.copyfileobj(source, target)

if __name__ == "__main__":
    restore = SystemRestore()
    backups = restore.list_backups()
    
    if backups:
        choice = input("Enter backup number to restore (or 'q' to quit): ")
        if choice.isdigit() and 1 <= int(choice) <= len(backups):
            backup_path = backups[int(choice) - 1]
            restore.restore_from_backup(backup_path)
```

## Troubleshooting Deployment Issues

### Common Deployment Problems

#### 9.1 Python Import Errors
```cmd
# Verify Python path
python -c "import sys; print('\n'.join(sys.path))"

# Check installed packages
pip list

# Reinstall packages if needed
pip install --force-reinstall -r requirements.txt
```

#### 9.2 Tesseract Not Found
```cmd
# Check Tesseract installation
where tesseract

# Verify path in configuration
python -c "
import json
with open('config/cv_config.json') as f:
    config = json.load(f)
print('Tesseract path:', config['ocr_settings']['tesseract_path'])
"
```

#### 9.3 Permission Issues
```cmd
# Check file permissions
icacls C:\VBS_OpenCV_Automation

# Fix permissions if needed
icacls C:\VBS_OpenCV_Automation /grant:r "%USERNAME%:(OI)(CI)F"
```

#### 9.4 Service Installation Issues
```cmd
# Check if service is installed
sc query VBSOpenCVAutomation

# Remove and reinstall service
python vbs_opencv_service.py remove
python vbs_opencv_service.py install
```

### Deployment Validation

#### 9.5 Post-Deployment Testing
Create `validate_deployment.py`:

```python
#!/usr/bin/env python3
"""
Deployment Validation Script
"""

import os
import sys
import json
import subprocess
from pathlib import Path

def validate_python_environment():
    """Validate Python environment"""
    print("Validating Python environment...")
    
    # Check Python version
    if sys.version_info < (3, 11):
        print("❌ Python 3.11+ required")
        return False
    
    print(f"✅ Python {sys.version}")
    
    # Check required packages
    required_packages = [
        'cv2', 'pytesseract', 'numpy', 'PIL', 
        'win32gui', 'win32api', 'psutil'
    ]
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package}")
        except ImportError:
            print(f"❌ {package} not found")
            return False
    
    return True

def validate_tesseract():
    """Validate Tesseract installation"""
    print("\nValidating Tesseract OCR...")
    
    try:
        result = subprocess.run(['tesseract', '--version'], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.split('\n')[0]
            print(f"✅ {version}")
            return True
        else:
            print("❌ Tesseract not accessible")
            return False
    except FileNotFoundError:
        print("❌ Tesseract not found in PATH")
        return False

def validate_configuration():
    """Validate configuration files"""
    print("\nValidating configuration...")
    
    config_file = Path("config/cv_config.json")
    if not config_file.exists():
        print("❌ Configuration file not found")
        return False
    
    try:
        with open(config_file) as f:
            config = json.load(f)
        
        # Check required sections
        required_sections = [
            'ocr_settings', 'template_matching', 
            'smart_automation', 'vbs_specific'
        ]
        
        for section in required_sections:
            if section in config:
                print(f"✅ {section}")
            else:
                print(f"❌ {section} missing")
                return False
        
        return True
        
    except json.JSONDecodeError:
        print("❌ Invalid JSON in configuration file")
        return False

def validate_directories():
    """Validate required directories"""
    print("\nValidating directories...")
    
    required_dirs = [
        "vbs/cv_services",
        "vbs/templates",
        "config",
        "EHC_Logs"
    ]
    
    all_exist = True
    for directory in required_dirs:
        if Path(directory).exists():
            print(f"✅ {directory}")
        else:
            print(f"❌ {directory} missing")
            all_exist = False
    
    return all_exist

def validate_cv_services():
    """Validate computer vision services"""
    print("\nValidating CV services...")
    
    try:
        from cv_services.ocr_service import OCRService
        from cv_services.template_service import TemplateService
        from cv_services.smart_engine import SmartAutomationEngine
        
        # Test OCR service
        ocr = OCRService()
        print("✅ OCR Service")
        
        # Test template service
        template = TemplateService()
        print("✅ Template Service")
        
        # Test smart engine
        engine = SmartAutomationEngine()
        print("✅ Smart Engine")
        
        return True
        
    except Exception as e:
        print(f"❌ CV services error: {e}")
        return False

def main():
    """Run all validation tests"""
    print("VBS OpenCV Deployment Validation")
    print("=" * 40)
    
    tests = [
        validate_python_environment,
        validate_tesseract,
        validate_configuration,
        validate_directories,
        validate_cv_services
    ]
    
    passed = 0
    total = len(tests)
    
    for test in tests:
        if test():
            passed += 1
    
    print(f"\nValidation Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 Deployment validation successful!")
        return True
    else:
        print("❌ Deployment validation failed!")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
```

## Maintenance Schedule

### Daily Tasks
- Check service status
- Review error logs
- Monitor system performance

### Weekly Tasks
- Rotate log files
- Update template images if needed
- Review performance statistics

### Monthly Tasks
- Create system backup
- Update configuration if needed
- Performance optimization review

### Quarterly Tasks
- Full system validation
- Dependency updates
- Security review

## Support and Documentation

### Getting Help
1. Check the troubleshooting guide: `docs/Troubleshooting_Guide.md`
2. Review log files in `EHC_Logs/`
3. Run the validation script: `python validate_deployment.py`
4. Check configuration: `config/cv_config.json`

### Documentation Files
- Main guide: `docs/VBS_OpenCV_Modernization_Guide.md`
- Performance tuning: `docs/Performance_Tuning_Guide.md`
- Template management: `docs/Template_Management_Guide.md`
- Troubleshooting: `docs/Troubleshooting_Guide.md`

This deployment guide provides comprehensive instructions for successfully deploying the VBS OpenCV modernization system in production environments.