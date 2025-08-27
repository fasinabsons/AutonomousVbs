#!/usr/bin/env python3
"""
Simple VBS Manager
For closing VBS software (AbsonsItERP.exe) cleanly
"""

import sys
import time
import psutil
import subprocess
import win32gui
import win32process
import logging
from datetime import datetime

class VBSManager:
    """Simple VBS application manager"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        
        # VBS process identifiers
        self.vbs_process_names = [
            "AbsonsItERP.exe",
            "absonsiterp.exe"
        ]
        
        # VBS window keywords
        self.vbs_window_keywords = [
            "absons",
            "arabian",
            "moonflower",
            "erp"
        ]
        
    def _setup_logging(self):
        """Simple logging setup"""
        logger = logging.getLogger("VBSManager")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def find_vbs_processes(self):
        """Find all VBS processes"""
        vbs_processes = []
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_info = proc.info
                    if proc_info['name'] and proc_info['exe']:
                        exe_name = proc_info['name'].lower()
                        exe_path = proc_info['exe'].lower()
                        
                        # Check if it's a VBS process
                        for vbs_name in self.vbs_process_names:
                            if vbs_name.lower() in exe_name or vbs_name.lower() in exe_path:
                                vbs_processes.append({
                                    'pid': proc_info['pid'],
                                    'name': proc_info['name'],
                                    'exe': proc_info['exe']
                                })
                                break
                                
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error finding VBS processes: {e}")
        
        return vbs_processes
    
    def find_vbs_windows(self):
        """Find all VBS windows"""
        vbs_windows = []
        
        def enum_callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd).lower()
                    
                    # Check if it's a VBS window
                    for keyword in self.vbs_window_keywords:
                        if keyword in title:
                            # Get process ID
                            _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                            windows.append({
                                'hwnd': hwnd,
                                'title': title,
                                'pid': process_id
                            })
                            break
                            
            except Exception:
                pass
            return True
        
        try:
            win32gui.EnumWindows(enum_callback, vbs_windows)
        except Exception as e:
            self.logger.error(f"Error finding VBS windows: {e}")
        
        return vbs_windows
    
    def close_vbs_windows_gracefully(self):
        """Close VBS windows gracefully first"""
        vbs_windows = self.find_vbs_windows()
        closed_count = 0
        
        if not vbs_windows:
            self.logger.info("No VBS windows found")
            return 0
        
        self.logger.info(f"Found {len(vbs_windows)} VBS windows")
        
        for window in vbs_windows:
            try:
                self.logger.info(f"Closing window: {window['title']} (PID: {window['pid']})")
                
                # Try to close window gracefully
                win32gui.PostMessage(window['hwnd'], win32gui.WM_CLOSE, 0, 0)
                time.sleep(1)
                
                # Check if window is still there
                if not win32gui.IsWindow(window['hwnd']) or not win32gui.IsWindowVisible(window['hwnd']):
                    self.logger.info(f"✅ Window closed gracefully")
                    closed_count += 1
                else:
                    self.logger.warning(f"⚠️ Window still open after close message")
                    
            except Exception as e:
                self.logger.error(f"Error closing window: {e}")
        
        return closed_count
    
    def terminate_vbs_processes(self):
        """Terminate VBS processes"""
        vbs_processes = self.find_vbs_processes()
        terminated_count = 0
        
        if not vbs_processes:
            self.logger.info("No VBS processes found")
            return 0
        
        self.logger.info(f"Found {len(vbs_processes)} VBS processes")
        
        for proc_info in vbs_processes:
            try:
                self.logger.info(f"Terminating process: {proc_info['name']} (PID: {proc_info['pid']})")
                
                process = psutil.Process(proc_info['pid'])
                
                # Try graceful termination first
                process.terminate()
                
                # Wait up to 5 seconds for graceful termination
                try:
                    process.wait(timeout=5)
                    self.logger.info(f"✅ Process terminated gracefully")
                    terminated_count += 1
                except psutil.TimeoutExpired:
                    # Force kill if graceful termination failed
                    self.logger.warning(f"⚠️ Graceful termination timeout, force killing...")
                    process.kill()
                    process.wait(timeout=3)
                    self.logger.info(f"✅ Process force killed")
                    terminated_count += 1
                    
            except psutil.NoSuchProcess:
                self.logger.info(f"✅ Process already gone")
                terminated_count += 1
            except Exception as e:
                self.logger.error(f"Error terminating process: {e}")
        
        return terminated_count
    
    def close_vbs_completely(self):
        """Close VBS application completely (windows + processes)"""
        self.logger.info("🛑 Closing VBS application completely...")
        print("🛑 Closing VBS application completely...")
        
        # Step 1: Close windows gracefully
        print("Step 1: Closing VBS windows gracefully...")
        windows_closed = self.close_vbs_windows_gracefully()
        print(f"✅ Closed {windows_closed} VBS windows")
        
        # Wait a moment for processes to exit naturally
        time.sleep(3)
        
        # Step 2: Terminate any remaining processes
        print("Step 2: Terminating any remaining VBS processes...")
        processes_terminated = self.terminate_vbs_processes()
        print(f"✅ Terminated {processes_terminated} VBS processes")
        
        # Step 3: Final verification
        time.sleep(2)
        remaining_processes = self.find_vbs_processes()
        remaining_windows = self.find_vbs_windows()
        
        if not remaining_processes and not remaining_windows:
            self.logger.info("✅ VBS application closed completely")
            print("✅ VBS application closed completely")
            return True
        else:
            self.logger.warning(f"⚠️ Some VBS components may still be running: {len(remaining_processes)} processes, {len(remaining_windows)} windows")
            print(f"⚠️ Some VBS components may still be running: {len(remaining_processes)} processes, {len(remaining_windows)} windows")
            return False
    
    def check_vbs_status(self):
        """Check if VBS is currently running"""
        processes = self.find_vbs_processes()
        windows = self.find_vbs_windows()
        
        status = {
            'running': len(processes) > 0 or len(windows) > 0,
            'processes': len(processes),
            'windows': len(windows),
            'process_list': processes,
            'window_list': windows
        }
        
        return status

def main():
    """Main function for command line usage"""
    vbs_manager = VBSManager()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        
        if command == "close":
            print("🛑 Closing VBS application...")
            result = vbs_manager.close_vbs_completely()
            sys.exit(0 if result else 1)
            
        elif command == "status":
            print("📊 Checking VBS status...")
            status = vbs_manager.check_vbs_status()
            
            if status['running']:
                print(f"🟢 VBS is running:")
                print(f"  Processes: {status['processes']}")
                print(f"  Windows: {status['windows']}")
                
                if status['process_list']:
                    print("  Process details:")
                    for proc in status['process_list']:
                        print(f"    - {proc['name']} (PID: {proc['pid']})")
                
                if status['window_list']:
                    print("  Window details:")
                    for win in status['window_list']:
                        print(f"    - {win['title']} (PID: {win['pid']})")
            else:
                print("🔴 VBS is not running")
            
            sys.exit(0)
            
        elif command == "kill":
            print("💀 Force killing all VBS processes...")
            terminated = vbs_manager.terminate_vbs_processes()
            print(f"✅ Terminated {terminated} processes")
            sys.exit(0)
            
        else:
            print("❌ Unknown command:", command)
            print("Usage: python vbs_manager.py [close|status|kill]")
            sys.exit(1)
    else:
        print("VBS Manager Utility")
        print("Manages AbsonsItERP.exe VBS application")
        print("")
        print("Usage: python vbs_manager.py [command]")
        print("Commands:")
        print("  close   - Close VBS application gracefully (windows + processes)")
        print("  status  - Check if VBS is currently running")
        print("  kill    - Force kill all VBS processes")
        print("")
        print("Examples:")
        print("  python vbs_manager.py close")
        print("  python vbs_manager.py status")
        print("  python vbs_manager.py kill")

if __name__ == "__main__":
    main() 