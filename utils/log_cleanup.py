#!/usr/bin/env python3
"""
Log Cleanup Utility for MoonFlower System
Centralizes all logs to EHC_Logs and cleans old logs (3+ days)
"""

import os
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path
import logging

class LogCleanup:
    """Centralized log management and cleanup"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.ehc_logs_dir = self.project_root / "EHC_Logs"
        self.today_folder = datetime.now().strftime("%d%b").lower()
        self.today_logs_dir = self.ehc_logs_dir / self.today_folder
        
        # Ensure today's log directory exists
        self.today_logs_dir.mkdir(parents=True, exist_ok=True)
        
        # Setup logging
        self.logger = self._setup_logging()
        
    def _setup_logging(self):
        """Setup logging for cleanup operations"""
        logger = logging.getLogger("LogCleanup")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # File handler in today's logs
            log_file = self.today_logs_dir / f"log_cleanup_{datetime.now().strftime('%Y%m%d')}.log"
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger
    
    def move_logs_to_ehc(self):
        """Move all scattered logs to EHC_Logs/today folder"""
        self.logger.info("Starting log centralization...")
        
        # Directories to check for logs
        search_dirs = [
            self.project_root,
            self.project_root / "AUTONOMOUS_LOGS",
            self.project_root / "email",
            self.project_root / "excel", 
            self.project_root / "wifi",
            self.project_root / "vbs",
            self.project_root / "utils"
        ]
        
        moved_count = 0
        
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
                
            # Find all log files
            for log_file in search_dir.rglob("*.log"):
                # Skip if already in EHC_Logs
                if "EHC_Logs" in str(log_file):
                    continue
                    
                try:
                    # Move to today's logs directory
                    destination = self.today_logs_dir / log_file.name
                    
                    # Handle duplicate names
                    counter = 1
                    while destination.exists():
                        stem = log_file.stem
                        suffix = log_file.suffix
                        destination = self.today_logs_dir / f"{stem}_{counter}{suffix}"
                        counter += 1
                    
                    shutil.move(str(log_file), str(destination))
                    self.logger.info(f"Moved: {log_file.name} -> {destination}")
                    moved_count += 1
                    
                except Exception as e:
                    self.logger.error(f"Failed to move {log_file}: {e}")
        
        self.logger.info(f"Centralization complete: {moved_count} logs moved to {self.today_logs_dir}")
    
    def cleanup_old_logs(self, days_to_keep=3):
        """Remove log folders older than specified days"""
        self.logger.info(f"Cleaning logs older than {days_to_keep} days...")
        
        cutoff_date = datetime.now() - timedelta(days=days_to_keep)
        removed_count = 0
        
        # Clean EHC_Logs directory
        for item in self.ehc_logs_dir.iterdir():
            if item.is_dir():
                try:
                    # Parse folder name (format: DDmon, e.g., 25jul)
                    folder_date = self._parse_folder_date(item.name)
                    
                    if folder_date and folder_date < cutoff_date:
                        shutil.rmtree(item)
                        self.logger.info(f"Removed old log folder: {item.name}")
                        removed_count += 1
                        
                except Exception as e:
                    self.logger.warning(f"Could not process folder {item.name}: {e}")
        
        # Clean individual log files in EHC_Logs root
        for log_file in self.ehc_logs_dir.glob("*.log"):
            try:
                file_time = datetime.fromtimestamp(log_file.stat().st_mtime)
                if file_time < cutoff_date:
                    log_file.unlink()
                    self.logger.info(f"Removed old log file: {log_file.name}")
                    removed_count += 1
            except Exception as e:
                self.logger.warning(f"Could not remove {log_file.name}: {e}")
        
        self.logger.info(f"Cleanup complete: {removed_count} items removed")
    
    def _parse_folder_date(self, folder_name):
        """Parse folder name to datetime (format: DDmon)"""
        try:
            # Extract day and month
            if len(folder_name) < 4:
                return None
                
            day_part = folder_name[:2]
            month_part = folder_name[2:]
            
            # Month mapping
            month_map = {
                'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
            }
            
            if month_part not in month_map:
                return None
                
            day = int(day_part)
            month = month_map[month_part]
            year = datetime.now().year
            
            return datetime(year, month, day)
            
        except (ValueError, KeyError):
            return None
    
    def get_log_summary(self):
        """Get summary of current log structure"""
        summary = {
            "total_folders": 0,
            "total_logs": 0,
            "today_logs": 0,
            "scattered_logs": 0
        }
        
        # Count EHC_Logs folders and files
        for item in self.ehc_logs_dir.iterdir():
            if item.is_dir():
                summary["total_folders"] += 1
                log_count = len(list(item.glob("*.log")))
                summary["total_logs"] += log_count
                
                if item.name == self.today_folder:
                    summary["today_logs"] = log_count
        
        # Count scattered logs
        search_dirs = [
            self.project_root / "AUTONOMOUS_LOGS",
            self.project_root / "email",
            self.project_root / "excel",
            self.project_root / "wifi", 
            self.project_root / "vbs",
            self.project_root / "utils"
        ]
        
        for search_dir in search_dirs:
            if search_dir.exists():
                summary["scattered_logs"] += len(list(search_dir.rglob("*.log")))
        
        return summary
    
    def run_full_cleanup(self):
        """Run complete log management process"""
        self.logger.info("=" * 60)
        self.logger.info("MOONFLOWER LOG CLEANUP - STARTING")
        self.logger.info("=" * 60)
        
        # Get initial summary
        initial_summary = self.get_log_summary()
        self.logger.info(f"Initial state: {initial_summary['scattered_logs']} scattered logs")
        
        # Move scattered logs to EHC_Logs
        self.move_logs_to_ehc()
        
        # Clean old logs (3+ days)
        self.cleanup_old_logs(days_to_keep=3)
        
        # Get final summary
        final_summary = self.get_log_summary()
        
        self.logger.info("=" * 60)
        self.logger.info("CLEANUP COMPLETE")
        self.logger.info(f"Total log folders: {final_summary['total_folders']}")
        self.logger.info(f"Total logs in EHC_Logs: {final_summary['total_logs']}")
        self.logger.info(f"Today's logs: {final_summary['today_logs']}")
        self.logger.info(f"Remaining scattered logs: {final_summary['scattered_logs']}")
        self.logger.info("=" * 60)


def main():
    """Main cleanup function"""
    cleanup = LogCleanup()
    cleanup.run_full_cleanup()


if __name__ == "__main__":
    main() 