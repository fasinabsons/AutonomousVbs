"""
File Manager for MoonFlower WiFi Automation
Handles daily folder creation and file organization
Now uses centralized PathManager for consistency
"""

import os
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict
import logging

try:
    from .path_manager import PathManager
    PATHMANAGER_AVAILABLE = True
except ImportError:
    PATHMANAGER_AVAILABLE = False

class FileManager:
    """Manages file system operations and daily folder structure"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        if PATHMANAGER_AVAILABLE:
            # Use centralized PathManager
            self.path_manager = PathManager()
            self.logger.info("FileManager using centralized PathManager")
        else:
            # Fallback to legacy implementation
            self.logger.warning("PathManager not available, using legacy implementation")
            
            # Get project root directory (parent of utils directory)
            utils_dir = Path(__file__).parent  # utils directory
            self.project_root = utils_dir.parent    # main project directory
            
            self.base_folders = {
                'csv': self.project_root / 'EHC_Data',
                'excel': self.project_root / 'EHC_Data_Merge', 
                'pdf': self.project_root / 'EHC_Data_Pdf',
                'logs': self.project_root / 'EHC_Logs'
            }
            self.setup_base_directories()
    
    def get_date_folder(self) -> str:
        """Get date-based folder name in DDMmm format (e.g., 04jul, 12aug)"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.get_date_folder()
        else:
            # Legacy implementation
            now = datetime.now()
            day = f"{now.day:02d}"
            month = now.strftime("%b").lower()
            return f"{day}{month}"
    
    def setup_base_directories(self) -> None:
        """Create base directory structure"""
        if PATHMANAGER_AVAILABLE:
            # PathManager handles this automatically
            return
        
        # Legacy implementation
        try:
            for folder_type, folder_path in self.base_folders.items():
                folder_path.mkdir(exist_ok=True)
                self.logger.info(f"Base directory ensured: {folder_path}")
        except Exception as e:
            self.logger.error(f"Failed to create base directories: {e}")
    
    def setup_daily_folders(self) -> None:
        """Create daily folders for all data types"""
        if PATHMANAGER_AVAILABLE:
            # PathManager handles this automatically in get_*_directory methods
            directories = self.path_manager.get_all_directories()
            for dir_type, path in directories.items():
                self.logger.info(f"Daily folder ensured: {path}")
        else:
            # Legacy implementation
            try:
                date_folder = self.get_date_folder()
                
                for folder_type, base_folder in self.base_folders.items():
                    daily_path = base_folder / date_folder
                    daily_path.mkdir(parents=True, exist_ok=True)
                    self.logger.info(f"Daily folder ensured: {daily_path}")
                    
                    # Create debug subfolder for logs
                    if folder_type == 'logs':
                        debug_path = daily_path / 'debug_images'
                        debug_path.mkdir(exist_ok=True)
                        
            except Exception as e:
                self.logger.error(f"Failed to create daily folders: {e}")
    
    def get_csv_directory(self) -> Path:
        """Get today's CSV download directory"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.get_csv_directory()
        else:
            # Legacy implementation
            date_folder = self.get_date_folder()
            csv_path = self.base_folders['csv'] / date_folder
            csv_path.mkdir(parents=True, exist_ok=True)
            return csv_path
    
    def get_excel_directory(self) -> Path:
        """Get today's Excel output directory"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.get_excel_directory()
        else:
            # Legacy implementation
            date_folder = self.get_date_folder()
            excel_path = self.base_folders['excel'] / date_folder
            excel_path.mkdir(parents=True, exist_ok=True)
            return excel_path
    
    def get_pdf_directory(self) -> Path:
        """Get today's PDF output directory"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.get_pdf_directory()
        else:
            # Legacy implementation
            date_folder = self.get_date_folder()
            pdf_path = self.base_folders['pdf'] / date_folder
            pdf_path.mkdir(parents=True, exist_ok=True)
            return pdf_path
    
    def get_logs_directory(self) -> Path:
        """Get today's logs directory"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.get_logs_directory()
        else:
            # Legacy implementation
            date_folder = self.get_date_folder()
            logs_path = self.base_folders['logs'] / date_folder
            logs_path.mkdir(parents=True, exist_ok=True)
            return logs_path
    
    def get_debug_directory(self) -> Path:
        """Get today's debug images directory"""
        logs_dir = self.get_logs_directory()
        debug_path = logs_dir / 'debug_images'
        debug_path.mkdir(exist_ok=True)
        return debug_path
    
    def count_csv_files(self, directory=None):
        """Count CSV files in directory"""
        if PATHMANAGER_AVAILABLE and directory is None:
            return self.path_manager.count_csv_files()
        
        # Legacy/custom directory implementation
        if directory is None:
            directory = self.get_csv_directory()
        try:
            return len(list(Path(directory).glob("*.csv")))
        except:
            return 0
    
    def get_excel_filename(self) -> str:
        """Generate Excel filename with today's date"""
        today = datetime.now().strftime("%d%m%Y")
        return f"EHC_Upload_Mac_{today}.xls"
    
    def get_pdf_filename(self) -> str:
        """Generate PDF filename with today's date"""
        today = datetime.now().strftime("%d%m%Y")
        return f"moon flower active users_{today}.pdf"
    
    def cleanup_old_files(self, days_to_keep: int = 30) -> None:
        """Clean up files older than specified days"""
        try:
            cutoff_date = datetime.now().timestamp() - (days_to_keep * 24 * 60 * 60)
            
            for folder_type, base_folder in self.base_folders.items():
                base_path = Path(base_folder)
                if base_path.exists():
                    for item in base_path.rglob("*"):
                        if item.is_file() and item.stat().st_mtime < cutoff_date:
                            item.unlink()
                            self.logger.info(f"Cleaned up old file: {item}")
                            
        except Exception as e:
            self.logger.error(f"Failed to cleanup old files: {e}")
    
    def get_file_paths(self) -> dict:
        """Get all current file paths for today"""
        date_folder = self.get_date_folder()
        excel_filename = self.get_excel_filename()
        pdf_filename = self.get_pdf_filename()
        
        return {
            'csv_dir': self.get_csv_directory(),
            'excel_file': self.get_excel_directory() / excel_filename,
            'pdf_file': self.get_pdf_directory() / pdf_filename,
            'logs_dir': self.get_logs_directory(),
            'debug_dir': self.get_debug_directory(),
            'date_folder': date_folder
        }
    
    def validate_paths(self) -> Dict[str, bool]:
        """Validate that all required paths exist and are accessible"""
        if PATHMANAGER_AVAILABLE:
            return self.path_manager.validate_paths()
        
        # Legacy implementation
        validation_results = {}
        date_folder = self.get_date_folder()
        
        for folder_type, base_path in self.base_folders.items():
            try:
                daily_path = base_path / date_folder
                daily_path.mkdir(parents=True, exist_ok=True)
                test_file = daily_path / "test_write.tmp"
                test_file.write_text("test")
                test_file.unlink()
                validation_results[folder_type] = True
            except Exception:
                validation_results[folder_type] = False
        
        return validation_results