#!/usr/bin/env python3
"""
Standalone VBS Application Closer
Utility to safely close VBS software processes

Usage:
    python utils/close_vbs.py
    python utils/close_vbs.py --force
    python utils/close_vbs.py --wait-seconds 10
"""

import sys
import time
import psutil
import subprocess
import argparse
from pathlib import Path

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Simple logging setup - compatible with existing utils
import logging

def setup_logger(name='VBSCloser'):
    """Setup simple logger compatible with existing utils"""
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    return logger

class VBSCloser:
    """Robust VBS application closer with multiple strategies"""
    
    def __init__(self, logger=None):
        self.logger = logger or setup_logger('VBSCloser')
        
        # VBS process names to look for (case-insensitive)
        self.vbs_process_names = [
            'AbsonsItERP',
            'ArabianLive', 
            'Arabian',
            'VBS',
            'vbs',
            'Visual Basic',
            'VisualBasic',
            'dotnet',
            'WindowsApplication'
        ]
    
    def find_vbs_processes(self):
        """Find all running VBS-related processes"""
        vbs_processes = []
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'exe']):
                try:
                    proc_name = proc.info['name'] or ''
                    proc_exe = proc.info['exe'] or ''
                    
                    # Check if process name matches VBS patterns
                    for vbs_name in self.vbs_process_names:
                        if (vbs_name.lower() in proc_name.lower() or 
                            vbs_name.lower() in proc_exe.lower()):
                            vbs_processes.append({
                                'pid': proc.info['pid'],
                                'name': proc_name,
                                'exe': proc_exe,
                                'process': proc
                            })
                            break
                            
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error finding VBS processes: {e}")
            
        return vbs_processes
    
    def close_process_gracefully(self, process_info, timeout=10):
        """Attempt graceful process termination"""
        try:
            proc = process_info['process']
            pid = process_info['pid']
            name = process_info['name']
            
            self.logger.info(f"Attempting graceful close of {name} (PID: {pid})")
            
            # Try terminate first
            proc.terminate()
            
            # Wait for graceful shutdown
            try:
                proc.wait(timeout=timeout)
                self.logger.info(f"Process {name} (PID: {pid}) closed gracefully")
                return True
            except psutil.TimeoutExpired:
                self.logger.warning(f"Process {name} (PID: {pid}) did not respond to terminate")
                return False
                
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            self.logger.info(f"Process {name} (PID: {pid}) already closed or access denied")
            return True
        except Exception as e:
            self.logger.error(f"Error during graceful close of {name}: {e}")
            return False
    
    def force_kill_process(self, process_info):
        """Force kill a stubborn process"""
        try:
            proc = process_info['process']
            pid = process_info['pid']
            name = process_info['name']
            
            self.logger.warning(f"Force killing {name} (PID: {pid})")
            proc.kill()
            
            # Verify it's dead
            time.sleep(1)
            if not proc.is_running():
                self.logger.info(f"Process {name} (PID: {pid}) force killed successfully")
                return True
            else:
                self.logger.error(f"Failed to force kill {name} (PID: {pid})")
                return False
                
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            self.logger.info(f"Process {name} (PID: {pid}) already terminated")
            return True
        except Exception as e:
            self.logger.error(f"Error during force kill of {name}: {e}")
            return False
    
    def close_via_taskkill(self, process_name):
        """Use Windows taskkill as last resort"""
        try:
            self.logger.info(f"Using taskkill for {process_name}")
            
            # Try graceful first
            result = subprocess.run([
                'taskkill', '/F', '/IM', f'{process_name}.exe'
            ], capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                self.logger.info(f"taskkill successful for {process_name}")
                return True
            else:
                self.logger.warning(f"taskkill failed for {process_name}: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error(f"taskkill timeout for {process_name}")
            return False
        except Exception as e:
            self.logger.error(f"taskkill error for {process_name}: {e}")
            return False
    
    def close_all_vbs(self, force=False, wait_seconds=5):
        """Main method to close all VBS processes"""
        self.logger.info("Starting VBS closure process")
        
        # Find all VBS processes
        vbs_processes = self.find_vbs_processes()
        
        if not vbs_processes:
            self.logger.info("No VBS processes found")
            return True
        
        self.logger.info(f"Found {len(vbs_processes)} VBS processes to close")
        
        success_count = 0
        total_count = len(vbs_processes)
        
        # Strategy 1: Graceful termination
        if not force:
            for proc_info in vbs_processes[:]:  # Copy list to avoid modification issues
                if self.close_process_gracefully(proc_info, timeout=wait_seconds):
                    success_count += 1
                    vbs_processes.remove(proc_info)
        
        # Strategy 2: Force kill remaining processes
        if vbs_processes:
            self.logger.warning(f"{len(vbs_processes)} processes require force termination")
            for proc_info in vbs_processes[:]:
                if self.force_kill_process(proc_info):
                    success_count += 1
                    vbs_processes.remove(proc_info)
        
        # Strategy 3: taskkill for any remaining
        if vbs_processes:
            self.logger.warning(f"{len(vbs_processes)} processes require taskkill")
            unique_names = set(proc['name'] for proc in vbs_processes)
            for name in unique_names:
                if self.close_via_taskkill(name):
                    # Count remaining processes with this name as successful
                    remaining = [p for p in vbs_processes if p['name'] == name]
                    success_count += len(remaining)
        
        # Final verification
        time.sleep(2)
        remaining_processes = self.find_vbs_processes()
        
        if remaining_processes:
            self.logger.error(f"{len(remaining_processes)} VBS processes still running after closure attempts")
            for proc in remaining_processes:
                self.logger.error(f"  - {proc['name']} (PID: {proc['pid']})")
            return False
        else:
            self.logger.info(f"All VBS processes closed successfully ({success_count}/{total_count})")
            return True

def main():
    parser = argparse.ArgumentParser(description='Close VBS software processes')
    parser.add_argument('--force', action='store_true', 
                       help='Skip graceful termination and force kill immediately')
    parser.add_argument('--wait-seconds', type=int, default=5,
                       help='Seconds to wait for graceful termination (default: 5)')
    parser.add_argument('--quiet', action='store_true',
                       help='Suppress informational output')
    
    args = parser.parse_args()
    
    # Setup logging
    if args.quiet:
        logging.getLogger().setLevel(logging.WARNING)
    
    # Create closer and execute
    closer = VBSCloser()
    success = closer.close_all_vbs(force=args.force, wait_seconds=args.wait_seconds)
    
    # Exit with appropriate code
    sys.exit(0 if success else 1)

if __name__ == '__main__':
    main()
