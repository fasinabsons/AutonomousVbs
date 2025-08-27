#!/usr/bin/env python3
"""
Configuration Loader for MoonFlower CSV Download System
Reads settings from config.txt file
"""

import os
from pathlib import Path
from typing import Dict, List, Any
import logging

class ConfigLoader:
    """Loads and manages configuration settings"""
    
    def __init__(self, config_file: str = "config.txt"):
        self.config_file = Path(config_file)
        self.config = {}
        self.logger = self._setup_logging()
        self.load_config()
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for config loader"""
        logger = logging.getLogger("ConfigLoader")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        return logger
    
    def load_config(self):
        """Load configuration from config.txt file"""
        try:
            if not self.config_file.exists():
                self.logger.warning(f"Config file {self.config_file} not found, using defaults")
                self._load_defaults()
                return
            
            self.logger.info(f"Loading configuration from {self.config_file}")
            
            with open(self.config_file, 'r') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    
                    # Skip comments and empty lines
                    if not line or line.startswith('#'):
                        continue
                    
                    # Parse key=value pairs
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # Convert values to appropriate types
                        self.config[key] = self._convert_value(value)
                    else:
                        self.logger.warning(f"Invalid config line {line_num}: {line}")
            
            # Validate and set defaults for missing keys
            self._validate_config()
            
            self.logger.info(f"Configuration loaded successfully with {len(self.config)} settings")
            
        except Exception as e:
            self.logger.error(f"Failed to load configuration: {e}")
            self._load_defaults()
    
    def _convert_value(self, value: str) -> Any:
        """Convert string value to appropriate type"""
        # Boolean values
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        
        # Integer values
        try:
            return int(value)
        except ValueError:
            pass
        
        # Float values
        try:
            return float(value)
        except ValueError:
            pass
        
        # Comma-separated lists
        if ',' in value:
            return [item.strip() for item in value.split(',')]
        
        # String values
        return value
    
    def _load_defaults(self):
        """Load default configuration values"""
        self.config = {
            'CSV_BASE_DIR': 'EHC_Data',
            'EXCEL_BASE_DIR': 'EHC_Data_Merge',
            'PDF_BASE_DIR': 'EHC_Data_Pdf',
            'LOG_BASE_DIR': 'EHC_Logs',
            'MORNING_START_TIME': '08:00',
            'MORNING_END_TIME': '12:59',
            'MORNING_DOWNLOAD_TIMES': ['08:30', '09:30', '10:30', '11:30', '12:30'],
            'AFTERNOON_START_TIME': '13:00',
            'AFTERNOON_END_TIME': '17:00',
            'AFTERNOON_DOWNLOAD_TIMES': ['13:30', '14:30', '15:30', '16:30'],
            'NOTIFICATION_EMAIL': 'faseenm@gmail.com',
            'EMAIL_SUBJECT_PREFIX': '📊 CSV Download Complete',
            'CHECK_INTERVAL': 1,
            'MAX_RETRY_ATTEMPTS': 3,
            'DOWNLOAD_TIMEOUT': 300,
            'DEBUG_LOGGING': False,
            'BROWSER_INSPECTION_TIME': 10,
            'AUTO_INSTALL_DEPENDENCIES': True
        }
        self.logger.info("Using default configuration values")
    
    def _validate_config(self):
        """Validate configuration and set defaults for missing keys"""
        defaults = {
            'CSV_BASE_DIR': 'EHC_Data',
            'EXCEL_BASE_DIR': 'EHC_Data_Merge',
            'PDF_BASE_DIR': 'EHC_Data_Pdf',
            'LOG_BASE_DIR': 'EHC_Logs',
            'MORNING_START_TIME': '08:00',
            'MORNING_END_TIME': '12:59',
            'MORNING_DOWNLOAD_TIMES': ['08:30', '09:30', '10:30', '11:30', '12:30'],
            'AFTERNOON_START_TIME': '13:00',
            'AFTERNOON_END_TIME': '17:00',
            'AFTERNOON_DOWNLOAD_TIMES': ['13:30', '14:30', '15:30', '16:30'],
            'NOTIFICATION_EMAIL': 'faseenm@gmail.com',
            'EMAIL_SUBJECT_PREFIX': '📊 CSV Download Complete',
            'CHECK_INTERVAL': 1,
            'MAX_RETRY_ATTEMPTS': 3,
            'DOWNLOAD_TIMEOUT': 300,
            'DEBUG_LOGGING': False,
            'BROWSER_INSPECTION_TIME': 10,
            'AUTO_INSTALL_DEPENDENCIES': True
        }
        
        for key, default_value in defaults.items():
            if key not in self.config:
                self.config[key] = default_value
                self.logger.info(f"Using default value for {key}: {default_value}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value"""
        return self.config.get(key, default)
    
    def get_folder_paths(self, date_folder: str) -> Dict[str, str]:
        """Get all folder paths for a given date"""
        base_paths = {
            'csv_dir': self.get('CSV_BASE_DIR'),
            'excel_dir': self.get('EXCEL_BASE_DIR'),
            'pdf_dir': self.get('PDF_BASE_DIR'),
            'log_dir': self.get('LOG_BASE_DIR')
        }
        
        # Add date folder to each path
        folder_paths = {}
        for key, base_path in base_paths.items():
            folder_paths[key] = str(Path(base_path) / date_folder)
        
        return folder_paths
    
    def get_time_slots(self) -> Dict[str, Dict[str, Any]]:
        """Get time slot configuration"""
        return {
            "morning": {
                "start_time": self.get('MORNING_START_TIME'),
                "end_time": self.get('MORNING_END_TIME'),
                "download_times": self.get('MORNING_DOWNLOAD_TIMES')
            },
            "afternoon": {
                "start_time": self.get('AFTERNOON_START_TIME'),
                "end_time": self.get('AFTERNOON_END_TIME'),
                "download_times": self.get('AFTERNOON_DOWNLOAD_TIMES')
            }
        }
    
    def get_email_config(self) -> Dict[str, str]:
        """Get email configuration"""
        return {
            'recipient': self.get('NOTIFICATION_EMAIL'),
            'subject_prefix': self.get('EMAIL_SUBJECT_PREFIX')
        }
    
    def create_directories(self, date_folder: str) -> Dict[str, str]:
        """Create all required directories and return their paths"""
        folder_paths = self.get_folder_paths(date_folder)
        
        created_dirs = {}
        for key, path in folder_paths.items():
            try:
                Path(path).mkdir(parents=True, exist_ok=True)
                created_dirs[key] = str(Path(path).absolute())
                self.logger.info(f"Created/verified directory: {created_dirs[key]}")
            except Exception as e:
                self.logger.error(f"Failed to create directory {path}: {e}")
                created_dirs[key] = path
        
        return created_dirs
    
    def reload_config(self):
        """Reload configuration from file"""
        self.logger.info("Reloading configuration...")
        self.config.clear()
        self.load_config()

def main():
    """Test the configuration loader"""
    config = ConfigLoader()
    
    print("=== Configuration Test ===")
    print(f"CSV Base Dir: {config.get('CSV_BASE_DIR')}")
    print(f"Morning Times: {config.get('MORNING_DOWNLOAD_TIMES')}")
    print(f"Email: {config.get('NOTIFICATION_EMAIL')}")
    
    print("\n=== Time Slots ===")
    time_slots = config.get_time_slots()
    for session, settings in time_slots.items():
        print(f"{session.title()}: {settings['start_time']}-{settings['end_time']}")
        print(f"  Download times: {', '.join(settings['download_times'])}")
    
    print("\n=== Folder Paths ===")
    date_folder = "22jul"
    folder_paths = config.get_folder_paths(date_folder)
    for key, path in folder_paths.items():
        print(f"{key}: {path}")
    
    print("\n=== Creating Directories ===")
    created_dirs = config.create_directories(date_folder)
    for key, path in created_dirs.items():
        exists = Path(path).exists()
        print(f"{key}: {path} {'✅' if exists else '❌'}")

if __name__ == "__main__":
    main()