import os
import time
import shutil
import json
from datetime import datetime, timedelta
from pathlib import Path

class LogManager:
    def __init__(self, log_base_dir="EHC_Logs", archive_days=7, config_file="config/log_settings.json"):
        self.log_base_dir = Path(log_base_dir)
        self.config_file = config_file
        self.archive_days = self._load_config().get("retention_days", archive_days)
        self.archive_dir = self.log_base_dir / "archive"
        self.archive_dir.mkdir(parents=True, exist_ok=True)
    
    def _load_config(self):
        """Load log management configuration"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    return json.load(f)
            else:
                return self._create_default_config()
        except Exception as e:
            print(f"Config loading failed: {e}")
            return {"retention_days": 7, "archive_enabled": True}
    
    def _create_default_config(self):
        """Create default log configuration"""
        default_config = {
            "retention_days": 7,
            "archive_enabled": True,
            "max_log_size_mb": 50,
            "compress_old_logs": True,
            "log_categories": {
                "csv_download": True,
                "excel_merge": True,
                "vbs_automation": True,
                "email_reports": True,
                "error_logs": True
            }
        }
        
        try:
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w') as f:
                json.dump(default_config, f, indent=2)
        except Exception as e:
            print(f"Failed to save default config: {e}")
        
        return default_config
    
    def cleanup_old_logs(self):
        """Remove logs older than archive_days and organize current logs"""
        current_time = datetime.now()
        cutoff_time = current_time - timedelta(days=self.archive_days)
        
        if not self.log_base_dir.exists():
            return
        
        cleanup_stats = {
            "folders_processed": 0,
            "folders_archived": 0,
            "files_cleaned": 0,
            "space_freed_mb": 0
        }
        
        # Process date-based folders (e.g., 16jul, 17jul, etc.)
        for item in self.log_base_dir.iterdir():
            if item.is_dir() and item.name != "archive":
                cleanup_stats["folders_processed"] += 1
                folder_time = os.path.getmtime(item)
                folder_datetime = datetime.fromtimestamp(folder_time)
                
                if folder_datetime < cutoff_time:
                    # Calculate folder size before deletion
                    folder_size = sum(f.stat().st_size for f in item.rglob('*') if f.is_file())
                    cleanup_stats["space_freed_mb"] += folder_size / (1024 * 1024)
                    
                    # Archive before deletion
                    if self._load_config().get("archive_enabled", True):
                        archive_path = self.archive_dir / f"archived_{item.name}_{current_time.strftime('%Y%m%d_%H%M%S')}"
                        shutil.make_archive(str(archive_path), 'gztar', str(item))
                        cleanup_stats["folders_archived"] += 1
                    
                    # Remove original folder
                    shutil.rmtree(item)
                    cleanup_stats["files_cleaned"] += 1
                    print(f"Cleaned logs from {item.name} (freed {folder_size/(1024*1024):.1f}MB)")
        
        # Clean loose log files in root
        for log_file in self.log_base_dir.glob("*.log"):
            file_time = os.path.getmtime(log_file)
            file_datetime = datetime.fromtimestamp(file_time)
            
            if file_datetime < cutoff_time:
                file_size = log_file.stat().st_size
                cleanup_stats["space_freed_mb"] += file_size / (1024 * 1024)
                log_file.unlink()
                cleanup_stats["files_cleaned"] += 1
        
        print(f"Log cleanup completed: {cleanup_stats}")
    
    def organize_logs(self):
        """Organize logs by date and type"""
        today = datetime.now().strftime("%d%b").lower()
        today_dir = self.log_base_dir / today
        today_dir.mkdir(parents=True, exist_ok=True)
        
        # Move loose log files to today's directory
        for log_file in self.log_base_dir.glob("*.log"):
            if log_file.is_file():
                shutil.move(log_file, today_dir / log_file.name)
    
    def daily_maintenance(self):
        """Run daily log maintenance"""
        self.organize_logs()
        self.cleanup_old_logs()
        print(f"Log maintenance completed at {datetime.now()}")

if __name__ == "__main__":
    log_manager = LogManager()
    log_manager.daily_maintenance()