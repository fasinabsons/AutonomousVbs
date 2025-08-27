"""
Windows Service Wrapper for MoonFlower WiFi Automation
Uses pywin32 for background service operation
"""

import sys
import os
import time
import logging
import threading
from pathlib import Path
import traceback
from datetime import datetime
import subprocess
import json

try:
    import win32serviceutil
    import win32service
    import win32event
    import win32api
    import win32con
    import servicemanager
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False
    print("Warning: pywin32 not fully available. Service will run in standalone mode.")

from orchestrator import MasterOrchestrator
from utils.config_manager import ConfigManager


class MoonFlowerService:
    """Windows Service for MoonFlower WiFi Automation"""
    
    _svc_name_ = "MoonFlowerWiFiAutomation"
    _svc_display_name_ = "MoonFlower WiFi Automation Service"
    _svc_description_ = "Automated WiFi data collection and processing service for 365-day operation"
    
    def __init__(self, args=None):
        # Service state
        self.is_alive = True
        self.orchestrator = None
        self.service_thread = None
        
        # Setup logging
        self.setup_service_logging()
        
        # Configuration
        self.config = ConfigManager()
        
        if PYWIN32_AVAILABLE:
            try:
                self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_INFORMATION_TYPE,
                    servicemanager.PYS_SERVICE_STARTED,
                    (self._svc_name_, '')
                )
            except Exception as e:
                self.logger.warning(f"Windows service integration limited: {e}")
        else:
            self.hWaitStop = None
    
    def setup_service_logging(self):
        """Setup logging for the Windows service"""
        try:
            # Create logs directory if it doesn't exist
            logs_dir = Path("EHC_Logs")
            logs_dir.mkdir(exist_ok=True)
            
            # Setup service-specific logging
            log_file = logs_dir / "windows_service.log"
            
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file),
                    logging.StreamHandler()
                ]
            )
            
            self.logger = logging.getLogger(__name__)
            self.logger.info("Windows Service logging initialized")
            
        except Exception as e:
            # Fallback to basic logging
            self.logger = logging.getLogger(__name__)
            self.logger.error(f"Failed to setup service logging: {e}")
            if PYWIN32_AVAILABLE:
                try:
                    servicemanager.LogErrorMsg(f"Failed to setup service logging: {e}")
                except:
                    pass
    
    def SvcStop(self):
        """Stop the service"""
        try:
            self.logger.info("Service stop requested")
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            
            # Signal stop event
            win32event.SetEvent(self.hWaitStop)
            
            # Stop orchestrator
            self.is_alive = False
            if self.orchestrator:
                self.orchestrator.stop_continuous_service()
            
            # Wait for service thread to complete
            if self.service_thread and self.service_thread.is_alive():
                self.service_thread.join(timeout=30)
            
            self.logger.info("Service stopped successfully")
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, '')
            )
            
        except Exception as e:
            self.logger.error(f"Error stopping service: {e}")
            servicemanager.LogErrorMsg(f"Error stopping service: {e}")
    
    def SvcDoRun(self):
        """Main service execution"""
        try:
            self.logger.info("Starting MoonFlower WiFi Automation Service")
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, 'Service is starting...')
            )
            
            # Start service in separate thread
            self.service_thread = threading.Thread(target=self.run_service, daemon=False)
            self.service_thread.start()
            
            # Wait for stop signal
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
        except Exception as e:
            self.logger.error(f"Service execution failed: {e}")
            servicemanager.LogErrorMsg(f"Service execution failed: {e}")
            raise
    
    def run_service(self):
        """Run the main service logic"""
        try:
            self.logger.info("Initializing MoonFlower orchestrator...")
            
            # Initialize orchestrator
            self.orchestrator = MasterOrchestrator()
            
            # Start continuous service
            self.orchestrator.start_continuous_service()
            
            self.logger.info("MoonFlower service is now running continuously")
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, 'Service is now running')
            )
            
            # Service main loop
            while self.is_alive:
                try:
                    # Health check and monitoring
                    self.perform_health_check()
                    
                    # Sleep for 5 minutes between health checks
                    for _ in range(300):  # 5 minutes = 300 seconds
                        if not self.is_alive:
                            break
                        time.sleep(1)
                    
                except Exception as e:
                    self.logger.error(f"Service loop error: {e}")
                    # Continue running even if health check fails
                    time.sleep(60)  # Wait 1 minute before retrying
            
            self.logger.info("Service main loop ended")
            
        except Exception as e:
            self.logger.error(f"Service run failed: {e}")
            self.logger.error(traceback.format_exc())
            
            servicemanager.LogErrorMsg(f"Service run failed: {e}")
            
            # Try to restart service after error
            if self.is_alive:
                self.logger.info("Attempting service recovery...")
                time.sleep(30)  # Wait before recovery attempt
                try:
                    if self.orchestrator:
                        self.orchestrator.stop_continuous_service()
                    self.orchestrator = MasterOrchestrator()
                    self.orchestrator.start_continuous_service()
                    self.logger.info("Service recovery successful")
                except Exception as recovery_error:
                    self.logger.error(f"Service recovery failed: {recovery_error}")
    
    def perform_health_check(self):
        """Perform service health check"""
        try:
            if not self.orchestrator:
                self.logger.warning("Orchestrator not initialized")
                return
            
            # Get workflow status
            status = self.orchestrator.get_workflow_status()
            
            # Log current status
            self.logger.debug(f"Service health check - Workflow status: {status['workflow_status']}")
            
            # Check for failed tasks that need attention
            failed_tasks = [name for name, task_status in status['task_status'].items() 
                          if task_status == 'failed']
            
            if failed_tasks:
                self.logger.warning(f"Failed tasks detected: {failed_tasks}")
                
                # Log to Windows Event Log for critical failures
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_WARNING_TYPE,
                    servicemanager.PYS_SERVICE_STARTED,
                    (self._svc_name_, f'Failed tasks: {", ".join(failed_tasks)}')
                )
            
            # Check disk space
            self.check_disk_space()
            
            # Check log file sizes
            self.check_log_file_sizes()
            
        except Exception as e:
            self.logger.error(f"Health check failed: {e}")
    
    def check_disk_space(self):
        """Check available disk space"""
        try:
            import shutil
            
            # Check disk space for current directory
            total, used, free = shutil.disk_usage(".")
            free_gb = free // (1024**3)
            
            if free_gb < 1:  # Less than 1GB free
                self.logger.warning(f"Low disk space: {free_gb}GB free")
                servicemanager.LogMsg(
                    servicemanager.EVENTLOG_WARNING_TYPE,
                    servicemanager.PYS_SERVICE_STARTED,
                    (self._svc_name_, f'Low disk space: {free_gb}GB free')
                )
            
        except Exception as e:
            self.logger.error(f"Disk space check failed: {e}")
    
    def check_log_file_sizes(self):
        """Check and rotate log files if they're too large"""
        try:
            logs_dir = Path("EHC_Logs")
            if not logs_dir.exists():
                return
            
            # Check all log files
            for log_file in logs_dir.rglob("*.log"):
                if log_file.stat().st_size > 50 * 1024 * 1024:  # 50MB
                    self.logger.info(f"Rotating large log file: {log_file}")
                    
                    # Create backup
                    backup_file = log_file.with_suffix(f".log.{datetime.now().strftime('%Y%m%d_%H%M%S')}")
                    log_file.rename(backup_file)
                    
                    # Create new empty log file
                    log_file.touch()
            
        except Exception as e:
            self.logger.error(f"Log file size check failed: {e}")


def install_service():
    """Install the Windows service"""
    if not PYWIN32_AVAILABLE:
        print("pywin32 not available. Using service manager for installation.")
        from service_manager import ServiceManager
        manager = ServiceManager()
        result = manager.install_service()
        return result.get('success', False)
    
    try:
        print("Installing MoonFlower WiFi Automation Service...")
        
        # Install service
        win32serviceutil.InstallService(
            MoonFlowerService._svc_reg_class_,
            MoonFlowerService._svc_name_,
            MoonFlowerService._svc_display_name_,
            description=MoonFlowerService._svc_description_,
            startType=win32service.SERVICE_AUTO_START
        )
        
        print("Service installed successfully!")
        print(f"Service Name: {MoonFlowerService._svc_name_}")
        print(f"Display Name: {MoonFlowerService._svc_display_name_}")
        print("Start Type: Automatic")
        
        return True
        
    except Exception as e:
        print(f"Failed to install service: {e}")
        return False


def uninstall_service():
    """Uninstall the Windows service"""
    try:
        print("Uninstalling MoonFlower WiFi Automation Service...")
        
        # Stop service if running
        try:
            win32serviceutil.StopService(MoonFlowerService._svc_name_)
            print("Service stopped.")
        except:
            pass  # Service might not be running
        
        # Remove service
        win32serviceutil.RemoveService(MoonFlowerService._svc_name_)
        
        print("Service uninstalled successfully!")
        return True
        
    except Exception as e:
        print(f"Failed to uninstall service: {e}")
        return False


def start_service():
    """Start the Windows service"""
    try:
        print("Starting MoonFlower WiFi Automation Service...")
        win32serviceutil.StartService(MoonFlowerService._svc_name_)
        print("Service started successfully!")
        return True
    except Exception as e:
        print(f"Failed to start service: {e}")
        return False


def stop_service():
    """Stop the Windows service"""
    try:
        print("Stopping MoonFlower WiFi Automation Service...")
        win32serviceutil.StopService(MoonFlowerService._svc_name_)
        print("Service stopped successfully!")
        return True
    except Exception as e:
        print(f"Failed to stop service: {e}")
        return False


def get_service_status():
    """Get the current service status"""
    try:
        status = win32serviceutil.QueryServiceStatus(MoonFlowerService._svc_name_)
        status_map = {
            win32service.SERVICE_STOPPED: "Stopped",
            win32service.SERVICE_START_PENDING: "Start Pending",
            win32service.SERVICE_STOP_PENDING: "Stop Pending",
            win32service.SERVICE_RUNNING: "Running",
            win32service.SERVICE_CONTINUE_PENDING: "Continue Pending",
            win32service.SERVICE_PAUSE_PENDING: "Pause Pending",
            win32service.SERVICE_PAUSED: "Paused"
        }
        
        current_status = status_map.get(status[1], f"Unknown ({status[1]})")
        
        return {
            'installed': True,
            'status': current_status,
            'status_code': status[1],
            'service_type': status[0],
            'controls_accepted': status[2],
            'win32_exit_code': status[3],
            'service_exit_code': status[4],
            'check_point': status[5],
            'wait_hint': status[6]
        }
        
    except Exception as e:
        return {
            'installed': False,
            'error': str(e)
        }


if __name__ == '__main__':
    if len(sys.argv) == 1:
        # Run as service
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MoonFlowerService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Handle command line arguments
        command = sys.argv[1].lower()
        
        if command == 'install':
            install_service()
        elif command == 'uninstall':
            uninstall_service()
        elif command == 'start':
            start_service()
        elif command == 'stop':
            stop_service()
        elif command == 'status':
            status = get_service_status()
            print(f"Service Status: {status}")
        elif command == 'debug':
            # Run in debug mode (not as service)
            print("Running in debug mode...")
            service = MoonFlowerService([])
            service.run_service()
        else:
            print("Usage:")
            print("  python windows_service.py install   - Install service")
            print("  python windows_service.py uninstall - Uninstall service")
            print("  python windows_service.py start     - Start service")
            print("  python windows_service.py stop      - Stop service")
            print("  python windows_service.py status    - Check service status")
            print("  python windows_service.py debug     - Run in debug mode")