#!/usr/bin/env python3
"""
Centralized Path Manager for MoonFlower Automation System
Ensures all components use correct and consistent folder paths
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional, List
import logging

class PathManager:
    """Centralized path management for the entire automation system"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        
        # Get project root directory (parent of utils directory)
        self.utils_dir = Path(__file__).parent  # utils directory
        self.project_root = self.utils_dir.parent  # main project directory
        
        # Load configuration
        self.config = self._load_config()
        
        # Ensure all base directories exist
        self._ensure_base_directories()
        
        self.logger.info(f"PathManager initialized with project root: {self.project_root}")
    
    def _load_config(self) -> Dict:
        """Load path configuration from JSON file"""
        try:
            config_file = self.project_root / "config" / "paths_config.json"
            if config_file.exists():
                with open(config_file, 'r') as f:
                    config = json.load(f)
                self.logger.info("Loaded path configuration from paths_config.json")
                return config
            else:
                self.logger.warning("paths_config.json not found, using defaults")
                return self._get_default_config()
        except Exception as e:
            self.logger.error(f"Failed to load path config: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict:
        """Get default configuration if config file is not available"""
        return {
            "base_folders": {
                "csv_data": "EHC_Data",
                "excel_merge": "EHC_Data_Merge",
                "pdf_reports": "EHC_Data_Pdf",
                "logs": "EHC_Logs"
            },
            "date_format": "%d%b"
        }
    
    def _ensure_base_directories(self):
        """Ensure all base directories exist"""
        try:
            for folder_type, folder_name in self.config["base_folders"].items():
                folder_path = self.project_root / folder_name
                folder_path.mkdir(exist_ok=True)
                self.logger.debug(f"Base directory ensured: {folder_path}")
        except Exception as e:
            self.logger.error(f"Failed to create base directories: {e}")
    
    def get_date_folder(self, date_obj: Optional[datetime] = None) -> str:
        """Get date-based folder name in DDMmm format (e.g., 24jul, 25jul)"""
        if date_obj is None:
            date_obj = datetime.now()
        
        day = f"{date_obj.day:02d}"
        month = date_obj.strftime("%b").lower()
        return f"{day}{month}"
    
    def get_csv_directory(self, date_folder: Optional[str] = None) -> Path:
        """Get CSV data directory for specified date folder"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        csv_path = self.project_root / self.config["base_folders"]["csv_data"] / date_folder
        csv_path.mkdir(parents=True, exist_ok=True)
        return csv_path
    
    def get_excel_directory(self, date_folder: Optional[str] = None) -> Path:
        """Get Excel merge directory for specified date folder"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        excel_path = self.project_root / self.config["base_folders"]["excel_merge"] / date_folder
        excel_path.mkdir(parents=True, exist_ok=True)
        return excel_path
    
    def get_pdf_directory(self, date_folder: Optional[str] = None) -> Path:
        """Get PDF reports directory for specified date folder"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        pdf_path = self.project_root / self.config["base_folders"]["pdf_reports"] / date_folder
        pdf_path.mkdir(parents=True, exist_ok=True)
        return pdf_path
    
    def get_logs_directory(self, date_folder: Optional[str] = None) -> Path:
        """Get logs directory for specified date folder"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        logs_path = self.project_root / self.config["base_folders"]["logs"] / date_folder
        logs_path.mkdir(parents=True, exist_ok=True)
        return logs_path
    
    def get_all_directories(self, date_folder: Optional[str] = None) -> Dict[str, Path]:
        """Get all directories for a specified date folder"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        return {
            "csv": self.get_csv_directory(date_folder),
            "excel": self.get_excel_directory(date_folder),
            "pdf": self.get_pdf_directory(date_folder),
            "logs": self.get_logs_directory(date_folder)
        }
    
    def validate_paths(self, date_folder: Optional[str] = None) -> Dict[str, bool]:
        """Validate that all required paths exist and are accessible"""
        if date_folder is None:
            date_folder = self.get_date_folder()
        
        validation_results = {}
        directories = self.get_all_directories(date_folder)
        
        for dir_type, dir_path in directories.items():
            try:
                # Check if directory exists and is writable
                dir_path.mkdir(parents=True, exist_ok=True)
                test_file = dir_path / "test_write.tmp"
                test_file.write_text("test")
                test_file.unlink()
                validation_results[dir_type] = True
                self.logger.debug(f"{dir_type} directory validated: {dir_path}")
            except Exception as e:
                validation_results[dir_type] = False
                self.logger.error(f"{dir_type} directory validation failed: {dir_path} - {e}")
        
        return validation_results
    
    def count_csv_files(self, date_folder: Optional[str] = None) -> int:
        """Count CSV files in the specified date folder"""
        csv_dir = self.get_csv_directory(date_folder)
        try:
            return len(list(csv_dir.glob("*.csv")))
        except Exception:
            return 0
    
    def list_available_dates(self) -> Dict[str, List[str]]:
        """List all available date folders for each data type"""
        available_dates = {}
        
        for folder_type, folder_name in self.config["base_folders"].items():
            base_path = self.project_root / folder_name
            if base_path.exists():
                date_folders = [d.name for d in base_path.iterdir() if d.is_dir()]
                available_dates[folder_type] = sorted(date_folders)
            else:
                available_dates[folder_type] = []
        
        return available_dates
    
    def get_project_root(self) -> Path:
        """Get the project root directory"""
        return self.project_root
    
    def get_config(self) -> Dict:
        """Get the current configuration"""
        return self.config.copy()


def test_path_manager():
    """Test the PathManager functionality"""
    print("🧪 Testing PathManager...")
    
    pm = PathManager()
    
    # Test basic functionality
    print(f"📁 Project Root: {pm.get_project_root()}")
    print(f"📅 Today's Date Folder: {pm.get_date_folder()}")
    
    # Test directory creation
    directories = pm.get_all_directories()
    print(f"\n📂 All Directories for today:")
    for dir_type, path in directories.items():
        print(f"   {dir_type}: {path}")
    
    # Test validation
    validation = pm.validate_paths()
    print(f"\n✅ Validation Results:")
    for dir_type, is_valid in validation.items():
        status = "✅" if is_valid else "❌"
        print(f"   {status} {dir_type}")
    
    # Test CSV file counting
    csv_count = pm.count_csv_files()
    print(f"\n📄 CSV Files Today: {csv_count}")
    
    # Test available dates
    available_dates = pm.list_available_dates()
    print(f"\n📅 Available Date Folders:")
    for folder_type, dates in available_dates.items():
        print(f"   {folder_type}: {len(dates)} folders ({', '.join(dates[-3:]) if dates else 'none'})")
    
    print("\n🎉 PathManager test completed!")


if __name__ == "__main__":
    test_path_manager() 