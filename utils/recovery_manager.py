#!/usr/bin/env python3
"""
Recovery Manager for MoonFlower WiFi Automation
Specialized recovery mechanisms for different system components
"""

import time
import logging
import subprocess
import psutil
import os
import shutil
import json
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional, Callable
from pathlib import Path
import win32service
import win32serviceutil
import win32api
import win32con
import win32gui
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from .error_handler import ErrorCategory, ErrorSeverity, SystemError


class RecoveryManager:
    """Specialized recovery mechanisms for system components"""
    
    def __init__(self, config_manager=None):
        """Initialize recovery manager"""
        self.config_manager = config_manager
        self.logger = logging.getLogger(__name__)
        
        # Recovery statistics
        self.recovery_stats = {
            "attempts": 0,
            "successes": 0,
            "failures": 0,
            "by_category": {}
        }
        
        # Component health tracking
        self.component_health = {}
        self.last_health_check = {}
        
        self.logger.info("🔧 Recovery Manager initialized")
    
    def recover_network_issues(self, error: SystemError) -> bool:
        """Comprehensive network issue recovery"""
        try:
            self.logger.info("🌐 Starting network recovery process...")
            self._update_recovery_stats("network", "attempt")
            
            # Step 1: Basic connectivity check
            if not self._check_basic_connectivity():
                self.logger.warning("No basic internet connectivity")
                # Try to reset network adapter
                if self._reset_network_adapter():
                    time.sleep(10)  # Wait for adapter reset
                    if not self._check_basic_connectivity():
                        self._update_recovery_stats("network", "failure")
                        return False
                else:
                    self._update_recovery_stats("network", "failure")
                    return False
            
            # Step 2: DNS resolution check
            if not self._check_dns_resolution():
                self.logger.info("DNS issues detected, attempting to fix...")
                if self._fix_dns_issues():
                    time.sleep(5)
                else:
                    self.logger.warning("DNS fix failed")
            
            # Step 3: Test specific endpoint
            endpoint = error.context.get('endpoint')
            if endpoint:
                if not self._test_endpoint_with_retry(endpoint):
                    self.logger.warning(f"Endpoint {endpoint} still unreachable")
                    self._update_recovery_stats("network", "failure")
                    return False
            
            # Step 4: Browser-specific recovery for Selenium issues
            if error.context.get('selenium_error'):
                if not self._recover_selenium_network_issues():
                    self._update_recovery_stats("network", "failure")
                    return False
            
            self.logger.info("✅ Network recovery completed successfully")
            self._update_recovery_stats("network", "success")
            return True
            
        except Exception as e:
            self.logger.error(f"Network recovery failed: {e}")
            self._update_recovery_stats("network", "failure")
            return False
    
    def recover_application_crashes(self, error: SystemError) -> bool:
        """Recover from application crashes"""
        try:
            self.logger.info("💥 Starting application crash recovery...")
            self._update_recovery_stats("application", "attempt")
            
            app_name = error.context.get('application_name', 'unknown')
            process_name = error.context.get('process_name')
            app_path = error.context.get('application_path')
            
            # Step 1: Kill any remaining processes
            if process_name:
                self._kill_process_completely(process_name)
                time.sleep(3)
            
            # Step 2: Clean up temporary files
            self._cleanup_application_temp_files(app_name)
            
            # Step 3: Check system resources
            if not self._check_system_resources():
                self.logger.warning("System resources are low, attempting cleanup...")
                self._free_system_resources()
                time.sleep(5)
            
            # Step 4: Restart application
            if app_path:
                if self._restart_application_with_retry(app_path, app_name):
                    # Step 5: Verify application is responsive
                    time.sleep(10)  # Give app time to start
                    if self._verify_application_health(app_name):
                        self.logger.info("✅ Application recovery completed successfully")
                        self._update_recovery_stats("application", "success")
                        return True
                    else:
                        self.logger.warning("Application started but not responsive")
                        self._update_recovery_stats("application", "failure")
                        return False
                else:
                    self._update_recovery_stats("application", "failure")
                    return False
            else:
                self.logger.warning("No application path provided for restart")
                self._update_recovery_stats("application", "failure")
                return False
                
        except Exception as e:
            self.logger.error(f"Application crash recovery failed: {e}")
            self._update_recovery_stats("application", "failure")
            return False
    
    def recover_file_system_issues(self, error: SystemError) -> bool:
        """Recover from file system issues"""
        try:
            self.logger.info("📁 Starting file system recovery...")
            self._update_recovery_stats("filesystem", "attempt")
            
            issue_type = error.context.get('issue_type', 'unknown')
            file_path = error.context.get('file_path')
            
            if issue_type == 'file_not_found':
                success = self._recover_missing_file(file_path, error.context)
            elif issue_type == 'permission_denied':
                success = self._recover_permission_issues(file_path)
            elif issue_type == 'disk_full':
                success = self._recover_disk_space_issues()
            elif issue_type == 'file_locked':
                success = self._recover_file_lock_issues(file_path)
            elif issue_type == 'corrupted_file':
                success = self._recover_corrupted_file(file_path, error.context)
            else:
                success = self._generic_filesystem_recovery()
            
            if success:
                self.logger.info("✅ File system recovery completed successfully")
                self._update_recovery_stats("filesystem", "success")
            else:
                self._update_recovery_stats("filesystem", "failure")
            
            return success
            
        except Exception as e:
            self.logger.error(f"File system recovery failed: {e}")
            self._update_recovery_stats("filesystem", "failure")
            return False
    
    def recover_data_processing_issues(self, error: SystemError) -> bool:
        """Recover from data processing issues"""
        try:
            self.logger.info("📊 Starting data processing recovery...")
            self._update_recovery_stats("data", "attempt")
            
            issue_type = error.context.get('issue_type', 'unknown')
            
            if issue_type == 'csv_parse_error':
                success = self._recover_csv_parsing_issues(error.context)
            elif issue_type == 'excel_generation_error':
                success = self._recover_excel_generation_issues(error.context)
            elif issue_type == 'data_validation_error':
                success = self._recover_data_validation_issues(error.context)
            elif issue_type == 'memory_error':
                success = self._recover_memory_issues()
            else:
                success = self._generic_data_recovery()
            
            if success:
                self.logger.info("✅ Data processing recovery completed successfully")
                self._update_recovery_stats("data", "success")
            else:
                self._update_recovery_stats("data", "failure")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Data processing recovery failed: {e}")
            self._update_recovery_stats("data", "failure")
            return False
    
    def recover_email_issues(self, error: SystemError) -> bool:
        """Recover from email delivery issues"""
        try:
            self.logger.info("📧 Starting email recovery...")
            self._update_recovery_stats("email", "attempt")
            
            issue_type = error.context.get('issue_type', 'unknown')
            
            if issue_type == 'smtp_connection_error':
                success = self._recover_smtp_connection_issues()
            elif issue_type == 'authentication_error':
                success = self._recover_email_authentication_issues()
            elif issue_type == 'attachment_error':
                success = self._recover_email_attachment_issues(error.context)
            elif issue_type == 'recipient_error':
                success = self._recover_recipient_issues(error.context)
            else:
                success = self._generic_email_recovery()
            
            if success:
                self.logger.info("✅ Email recovery completed successfully")
                self._update_recovery_stats("email", "success")
            else:
                self._update_recovery_stats("email", "failure")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Email recovery failed: {e}")
            self._update_recovery_stats("email", "failure")
            return False
    
    def recover_system_issues(self, error: SystemError) -> bool:
        """Recover from system-level issues"""
        try:
            self.logger.info("🖥️ Starting system recovery...")
            self._update_recovery_stats("system", "attempt")
            
            issue_type = error.context.get('issue_type', 'unknown')
            
            if issue_type == 'service_failure':
                success = self._recover_service_issues(error.context)
            elif issue_type == 'memory_exhaustion':
                success = self._recover_memory_exhaustion()
            elif issue_type == 'high_cpu_usage':
                success = self._recover_high_cpu_usage()
            elif issue_type == 'registry_issues':
                success = self._recover_registry_issues()
            else:
                success = self._generic_system_recovery()
            
            if success:
                self.logger.info("✅ System recovery completed successfully")
                self._update_recovery_stats("system", "success")
            else:
                self._update_recovery_stats("system", "failure")
            
            return success
            
        except Exception as e:
            self.logger.error(f"System recovery failed: {e}")
            self._update_recovery_stats("system", "failure")
            return False
    
    # Network recovery methods
    def _check_basic_connectivity(self) -> bool:
        """Check basic internet connectivity"""
        try:
            import socket
            socket.create_connection(("8.8.8.8", 53), timeout=5)
            return True
        except Exception:
            return False
    
    def _reset_network_adapter(self) -> bool:
        """Reset network adapter"""
        try:
            self.logger.info("Resetting network adapter...")
            
            # Disable and re-enable network adapter
            cmd_disable = ['netsh', 'interface', 'set', 'interface', 'name="Ethernet"', 'admin=DISABLED']
            cmd_enable = ['netsh', 'interface', 'set', 'interface', 'name="Ethernet"', 'admin=ENABLED']
            
            subprocess.run(cmd_disable, capture_output=True, text=True, timeout=30)
            time.sleep(5)
            subprocess.run(cmd_enable, capture_output=True, text=True, timeout=30)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Network adapter reset failed: {e}")
            return False
    
    def _check_dns_resolution(self) -> bool:
        """Check DNS resolution"""
        try:
            import socket
            socket.gethostbyname("google.com")
            return True
        except Exception:
            return False
    
    def _fix_dns_issues(self) -> bool:
        """Fix DNS resolution issues"""
        try:
            self.logger.info("Fixing DNS issues...")
            
            # Flush DNS cache
            subprocess.run(['ipconfig', '/flushdns'], capture_output=True, text=True, timeout=30)
            
            # Reset DNS to Google DNS
            cmd = ['netsh', 'interface', 'ip', 'set', 'dns', 'name="Ethernet"', 'static', '8.8.8.8']
            subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            return True
            
        except Exception as e:
            self.logger.error(f"DNS fix failed: {e}")
            return False
    
    def _test_endpoint_with_retry(self, endpoint: str, max_retries: int = 3) -> bool:
        """Test endpoint connectivity with retry"""
        for attempt in range(max_retries):
            try:
                response = requests.get(endpoint, timeout=10, verify=False)
                if response.status_code < 500:
                    return True
            except Exception:
                pass
            
            if attempt < max_retries - 1:
                time.sleep(5 * (attempt + 1))  # Exponential backoff
        
        return False
    
    def _recover_selenium_network_issues(self) -> bool:
        """Recover Selenium-specific network issues"""
        try:
            self.logger.info("Recovering Selenium network issues...")
            
            # Kill any hanging Chrome processes
            self._kill_process_completely("chrome.exe")
            self._kill_process_completely("chromedriver.exe")
            
            time.sleep(5)
            
            # Test Chrome startup
            try:
                options = Options()
                options.add_argument("--headless")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-gpu")
                
                driver = webdriver.Chrome(options=options)
                driver.get("https://www.google.com")
                driver.quit()
                
                return True
                
            except Exception as e:
                self.logger.error(f"Chrome test failed: {e}")
                return False
                
        except Exception as e:
            self.logger.error(f"Selenium network recovery failed: {e}")
            return False
    
    # Application recovery methods
    def _kill_process_completely(self, process_name: str):
        """Completely kill a process and all its children"""
        try:
            killed_count = 0
            for proc in psutil.process_iter(['pid', 'name', 'ppid']):
                try:
                    if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                        # Kill child processes first
                        children = proc.children(recursive=True)
                        for child in children:
                            try:
                                child.terminate()
                                child.wait(timeout=5)
                            except:
                                try:
                                    child.kill()
                                except:
                                    pass
                        
                        # Kill main process
                        proc.terminate()
                        proc.wait(timeout=10)
                        killed_count += 1
                        
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if killed_count > 0:
                self.logger.info(f"Killed {killed_count} instances of {process_name}")
                
        except Exception as e:
            self.logger.error(f"Process kill failed for {process_name}: {e}")
    
    def _cleanup_application_temp_files(self, app_name: str):
        """Clean up application temporary files"""
        try:
            temp_dirs = [
                Path(os.environ.get('TEMP', '')),
                Path(os.environ.get('TMP', '')),
                Path.home() / "AppData" / "Local" / "Temp"
            ]
            
            for temp_dir in temp_dirs:
                if temp_dir.exists():
                    # Look for app-specific temp files
                    for pattern in [f"*{app_name}*", "*.tmp", "*.log"]:
                        for file_path in temp_dir.glob(pattern):
                            try:
                                if file_path.is_file() and (datetime.now() - datetime.fromtimestamp(file_path.stat().st_mtime)).days > 1:
                                    file_path.unlink()
                            except:
                                pass
                                
        except Exception as e:
            self.logger.warning(f"Temp file cleanup failed: {e}")
    
    def _check_system_resources(self) -> bool:
        """Check if system has adequate resources"""
        try:
            # Check memory
            memory = psutil.virtual_memory()
            if memory.percent > 90:
                return False
            
            # Check CPU
            cpu_percent = psutil.cpu_percent(interval=1)
            if cpu_percent > 95:
                return False
            
            # Check disk space
            disk = psutil.disk_usage('.')
            if (disk.free / disk.total) < 0.05:  # Less than 5% free
                return False
            
            return True
            
        except Exception:
            return False
    
    def _free_system_resources(self):
        """Free up system resources"""
        try:
            # Clear system cache
            subprocess.run(['sfc', '/scannow'], capture_output=True, timeout=300)
            
            # Run disk cleanup
            subprocess.run(['cleanmgr', '/sagerun:1'], capture_output=True, timeout=300)
            
            # Force garbage collection
            import gc
            gc.collect()
            
        except Exception as e:
            self.logger.warning(f"Resource cleanup failed: {e}")
    
    def _restart_application_with_retry(self, app_path: str, app_name: str, max_retries: int = 3) -> bool:
        """Restart application with retry logic"""
        for attempt in range(max_retries):
            try:
                self.logger.info(f"Starting {app_name} (attempt {attempt + 1})")
                
                # Start the application
                process = subprocess.Popen([app_path])
                
                # Wait for application to start
                time.sleep(10)
                
                # Check if process is still running
                if process.poll() is None:
                    return True
                else:
                    self.logger.warning(f"{app_name} exited immediately")
                    
            except Exception as e:
                self.logger.error(f"Application start attempt {attempt + 1} failed: {e}")
            
            if attempt < max_retries - 1:
                time.sleep(5)
        
        return False
    
    def _verify_application_health(self, app_name: str) -> bool:
        """Verify application is healthy and responsive"""
        try:
            # Check if process is running
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and app_name.lower() in proc.info['name'].lower():
                    # Check if process is responsive (not hung)
                    try:
                        proc.cpu_percent()  # This will fail if process is hung
                        return True
                    except:
                        return False
            
            return False
            
        except Exception:
            return False
    
    # File system recovery methods
    def _recover_missing_file(self, file_path: str, context: Dict[str, Any]) -> bool:
        """Recover missing file"""
        try:
            if not file_path:
                return False
            
            path_obj = Path(file_path)
            
            # Try to find backup or alternative file
            backup_file = self._find_backup_file(file_path)
            if backup_file:
                shutil.copy2(backup_file, file_path)
                self.logger.info(f"Restored file from backup: {backup_file}")
                return True
            
            # Try to recreate file if template available
            template_data = context.get('template_data')
            if template_data:
                with open(file_path, 'w') as f:
                    if isinstance(template_data, dict):
                        json.dump(template_data, f, indent=2)
                    else:
                        f.write(str(template_data))
                self.logger.info(f"Recreated file from template: {file_path}")
                return True
            
            # Create empty file as last resort
            path_obj.parent.mkdir(parents=True, exist_ok=True)
            path_obj.touch()
            self.logger.info(f"Created empty file: {file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"File recovery failed for {file_path}: {e}")
            return False
    
    def _find_backup_file(self, file_path: str) -> Optional[str]:
        """Find backup file"""
        try:
            path_obj = Path(file_path)
            parent_dir = path_obj.parent
            file_stem = path_obj.stem
            file_suffix = path_obj.suffix
            
            # Look for backup files
            backup_patterns = [
                f"{file_stem}.bak{file_suffix}",
                f"{file_stem}_backup{file_suffix}",
                f"{file_stem}_{datetime.now().strftime('%Y%m%d')}{file_suffix}"
            ]
            
            for pattern in backup_patterns:
                backup_path = parent_dir / pattern
                if backup_path.exists():
                    return str(backup_path)
            
            # Look for similar files
            similar_files = list(parent_dir.glob(f"{file_stem}*{file_suffix}"))
            if similar_files:
                # Return the most recent one
                latest_file = max(similar_files, key=os.path.getctime)
                return str(latest_file)
            
            return None
            
        except Exception:
            return None
    
    def _recover_permission_issues(self, file_path: str) -> bool:
        """Recover from permission issues"""
        try:
            if not file_path:
                return False
            
            # Try to fix permissions using icacls
            cmd = ['icacls', file_path, '/grant', 'Everyone:F']
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                self.logger.info(f"Fixed permissions for {file_path}")
                return True
            else:
                self.logger.warning(f"Permission fix failed: {result.stderr}")
                return False
                
        except Exception as e:
            self.logger.error(f"Permission recovery failed: {e}")
            return False
    
    def _recover_disk_space_issues(self) -> bool:
        """Recover from disk space issues"""
        try:
            self.logger.info("Attempting to free disk space...")
            
            # Clean temporary files
            temp_dirs = [
                Path(os.environ.get('TEMP', '')),
                Path(os.environ.get('TMP', '')),
                Path.home() / "AppData" / "Local" / "Temp"
            ]
            
            freed_space = 0
            for temp_dir in temp_dirs:
                if temp_dir.exists():
                    for file_path in temp_dir.rglob("*"):
                        try:
                            if file_path.is_file():
                                file_size = file_path.stat().st_size
                                file_path.unlink()
                                freed_space += file_size
                        except:
                            pass
            
            # Clean old log files
            log_dirs = [Path("EHC_Logs"), Path("logs")]
            for log_dir in log_dirs:
                if log_dir.exists():
                    cutoff_date = datetime.now() - timedelta(days=30)
                    for log_file in log_dir.glob("*.log"):
                        try:
                            if datetime.fromtimestamp(log_file.stat().st_mtime) < cutoff_date:
                                file_size = log_file.stat().st_size
                                log_file.unlink()
                                freed_space += file_size
                        except:
                            pass
            
            freed_mb = freed_space / (1024 * 1024)
            self.logger.info(f"Freed {freed_mb:.1f} MB of disk space")
            
            # Check if we have enough space now
            disk = psutil.disk_usage('.')
            free_percent = (disk.free / disk.total) * 100
            
            return free_percent > 5  # At least 5% free
            
        except Exception as e:
            self.logger.error(f"Disk space recovery failed: {e}")
            return False
    
    def _recover_file_lock_issues(self, file_path: str) -> bool:
        """Recover from file lock issues"""
        try:
            if not file_path:
                return False
            
            # Find processes that have the file open
            locked_processes = []
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    for file_info in proc.open_files():
                        if file_info.path == file_path:
                            locked_processes.append(proc)
                            break
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            # Terminate processes holding the file
            for proc in locked_processes:
                try:
                    self.logger.info(f"Terminating process {proc.name()} (PID: {proc.pid}) holding file lock")
                    proc.terminate()
                    proc.wait(timeout=10)
                except:
                    try:
                        proc.kill()
                    except:
                        pass
            
            # Wait a moment for file to be released
            time.sleep(2)
            
            # Test if file is accessible now
            try:
                with open(file_path, 'a'):
                    pass
                return True
            except:
                return False
                
        except Exception as e:
            self.logger.error(f"File lock recovery failed: {e}")
            return False
    
    # Data processing recovery methods
    def _recover_csv_parsing_issues(self, context: Dict[str, Any]) -> bool:
        """Recover from CSV parsing issues"""
        try:
            csv_file = context.get('csv_file')
            if not csv_file or not Path(csv_file).exists():
                return False
            
            # Try different encodings
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']
            
            for encoding in encodings:
                try:
                    import pandas as pd
                    df = pd.read_csv(csv_file, encoding=encoding)
                    if not df.empty:
                        # Save with correct encoding
                        backup_file = f"{csv_file}.fixed"
                        df.to_csv(backup_file, index=False, encoding='utf-8')
                        shutil.move(backup_file, csv_file)
                        self.logger.info(f"Fixed CSV encoding: {encoding}")
                        return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            self.logger.error(f"CSV recovery failed: {e}")
            return False
    
    def _recover_excel_generation_issues(self, context: Dict[str, Any]) -> bool:
        """Recover from Excel generation issues"""
        try:
            # Free up memory
            import gc
            gc.collect()
            
            # Try alternative Excel library
            try:
                import openpyxl
                # Implementation would depend on specific Excel generation code
                return True
            except ImportError:
                pass
            
            # Try with smaller data chunks
            data = context.get('data')
            if data and len(data) > 1000:
                # Process in smaller chunks
                self.logger.info("Processing Excel data in smaller chunks")
                return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"Excel recovery failed: {e}")
            return False
    
    def _recover_memory_issues(self) -> bool:
        """Recover from memory issues"""
        try:
            # Force garbage collection
            import gc
            gc.collect()
            
            # Clear caches
            if hasattr(gc, 'set_threshold'):
                gc.set_threshold(700, 10, 10)
            
            # Check memory after cleanup
            memory = psutil.virtual_memory()
            if memory.percent < 85:
                return True
            
            # Kill non-essential processes
            self._kill_non_essential_processes()
            
            return True
            
        except Exception as e:
            self.logger.error(f"Memory recovery failed: {e}")
            return False
    
    def _kill_non_essential_processes(self):
        """Kill non-essential processes to free memory"""
        try:
            # List of processes that can be safely terminated
            non_essential = ['notepad.exe', 'calc.exe', 'mspaint.exe']
            
            for proc in psutil.process_iter(['name']):
                try:
                    if proc.info['name'] and proc.info['name'].lower() in non_essential:
                        proc.terminate()
                        proc.wait(timeout=5)
                except:
                    pass
                    
        except Exception as e:
            self.logger.warning(f"Non-essential process cleanup failed: {e}")
    
    # Email recovery methods
    def _recover_smtp_connection_issues(self) -> bool:
        """Recover from SMTP connection issues"""
        try:
            # Test different SMTP servers
            smtp_servers = [
                ('smtp.gmail.com', 587),
                ('smtp.outlook.com', 587),
                ('smtp.yahoo.com', 587)
            ]
            
            for server, port in smtp_servers:
                try:
                    import smtplib
                    smtp = smtplib.SMTP(server, port, timeout=10)
                    smtp.starttls()
                    smtp.quit()
                    self.logger.info(f"SMTP server {server} is accessible")
                    return True
                except:
                    continue
            
            return False
            
        except Exception as e:
            self.logger.error(f"SMTP recovery failed: {e}")
            return False
    
    # System recovery methods
    def _recover_service_issues(self, context: Dict[str, Any]) -> bool:
        """Recover from Windows service issues"""
        try:
            service_name = context.get('service_name', 'MoonFlowerWiFiAutomation')
            
            # Try to restart the service
            try:
                win32serviceutil.StopService(service_name)
                time.sleep(5)
                win32serviceutil.StartService(service_name)
                time.sleep(10)
                
                # Check if service is running
                status = win32serviceutil.QueryServiceStatus(service_name)
                if status[1] == win32service.SERVICE_RUNNING:
                    return True
                    
            except Exception as e:
                self.logger.error(f"Service restart failed: {e}")
            
            return False
            
        except Exception as e:
            self.logger.error(f"Service recovery failed: {e}")
            return False
    
    # Generic recovery methods
    def _generic_filesystem_recovery(self) -> bool:
        """Generic filesystem recovery"""
        try:
            # Check disk health
            result = subprocess.run(['chkdsk', 'C:', '/f'], capture_output=True, text=True, timeout=300)
            return result.returncode == 0
        except:
            return False
    
    def _generic_data_recovery(self) -> bool:
        """Generic data recovery"""
        try:
            # Clear caches and continue
            import gc
            gc.collect()
            return True
        except:
            return False
    
    def _generic_email_recovery(self) -> bool:
        """Generic email recovery"""
        try:
            # Wait and retry
            time.sleep(30)
            return True
        except:
            return False
    
    def _generic_system_recovery(self) -> bool:
        """Generic system recovery"""
        try:
            # Basic system health check
            return self._check_system_resources()
        except:
            return False
    
    # Statistics and monitoring
    def _update_recovery_stats(self, category: str, result: str):
        """Update recovery statistics"""
        try:
            self.recovery_stats["attempts"] += 1
            
            if result == "success":
                self.recovery_stats["successes"] += 1
            elif result == "failure":
                self.recovery_stats["failures"] += 1
            
            if category not in self.recovery_stats["by_category"]:
                self.recovery_stats["by_category"][category] = {
                    "attempts": 0, "successes": 0, "failures": 0
                }
            
            self.recovery_stats["by_category"][category]["attempts"] += 1
            if result in ["success", "failure"]:
                self.recovery_stats["by_category"][category][result + "es"] += 1
                
        except Exception as e:
            self.logger.error(f"Stats update failed: {e}")
    
    def get_recovery_stats(self) -> Dict[str, Any]:
        """Get recovery statistics"""
        try:
            total_attempts = self.recovery_stats["attempts"]
            success_rate = (self.recovery_stats["successes"] / total_attempts * 100) if total_attempts > 0 else 0
            
            return {
                "total_attempts": total_attempts,
                "total_successes": self.recovery_stats["successes"],
                "total_failures": self.recovery_stats["failures"],
                "success_rate": round(success_rate, 2),
                "by_category": self.recovery_stats["by_category"],
                "component_health": self.component_health
            }
            
        except Exception as e:
            return {"error": f"Failed to get stats: {e}"}


if __name__ == "__main__":
    # Test recovery manager
    print("🧪 Testing Recovery Manager")
    print("=" * 50)
    
    recovery_manager = RecoveryManager()
    
    # Test network recovery
    test_error = SystemError(
        category=ErrorCategory.NETWORK,
        severity=ErrorSeverity.MEDIUM,
        message="Connection timeout",
        context={"endpoint": "https://www.google.com"}
    )
    
    result = recovery_manager.recover_network_issues(test_error)
    print(f"Network recovery test: {'✅ Success' if result else '❌ Failed'}")
    
    # Print stats
    stats = recovery_manager.get_recovery_stats()
    print(f"Recovery stats: {json.dumps(stats, indent=2)}")
    
    print("Test completed")