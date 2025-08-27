#!/usr/bin/env python3
"""
Universal Path Manager for MoonFlower Automation System
Handles all path resolution issues and works from any directory location
"""

import os
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, Optional

class UniversalPathManager:
    """Resolves all path issues for the automation system"""
    
    def __init__(self):
        """Initialize with automatic path detection"""
        self.current_working_dir = Path.cwd()
        self.script_dir = Path(__file__).parent
        
        # Detect project root intelligently
        self.project_root = self._detect_project_root()
        
        # Date folder for daily operations
        self.date_folder = datetime.now().strftime("%d%b").lower()
        
        print(f"📁 UniversalPathManager initialized:")
        print(f"   Current Working Dir: {self.current_working_dir}")
        print(f"   Script Dir: {self.script_dir}")
        print(f"   Project Root: {self.project_root}")
        print(f"   Date Folder: {self.date_folder}")
    
    def _detect_project_root(self) -> Path:
        """Use FIXED project root path for C:\\Users\\Lenovo\\Documents\\Automate2\\Automata2"""
        
        # FIXED PATH for the target system (corrected)
        fixed_root = Path("C:/Users/Lenovo/Documents/Automate2/Automata2")
        
        # If the fixed path exists, use it
        if fixed_root.exists():
            print(f"✅ Using FIXED project root: {fixed_root}")
            return fixed_root
        
        # If fixed path doesn't exist, try to create it
        try:
            fixed_root.mkdir(parents=True, exist_ok=True)
            print(f"✅ Created and using FIXED project root: {fixed_root}")
            return fixed_root
        except Exception as e:
            print(f"⚠️ Could not create fixed root, using fallback detection: {e}")
        
        # Fallback: Original detection logic
        indicators = ['3_VBS_Upload.bat', '4_VBS_Report.bat', 'config', 'vbs', 'wifi', 'email']
        search_dir = self.current_working_dir
        
        for _ in range(5):  # Max 5 levels up
            found_indicators = 0
            for indicator in indicators:
                if (search_dir / indicator).exists():
                    found_indicators += 1
            
            if found_indicators >= 3:
                return search_dir
            
            search_dir = search_dir.parent
        
        print(f"⚠️ Could not detect project root, using current directory: {self.current_working_dir}")
        return self.current_working_dir
    
    def get_paths(self) -> Dict[str, Path]:
        """Get all important paths for the automation system"""
        
        return {
            # Base directories
            'project_root': self.project_root,
            'vbs_dir': self.project_root / 'vbs',
            'wifi_dir': self.project_root / 'wifi',
            'email_dir': self.project_root / 'email',
            'excel_dir': self.project_root / 'excel',
            'config_dir': self.project_root / 'config',
            'utils_dir': self.project_root / 'utils',
            'images_dir': self.project_root / 'Images',
            
            # Data directories (with date folder)
            'csv_data': self.project_root / 'EHC_Data' / self.date_folder,
            'excel_merge': self.project_root / 'EHC_Data_Merge' / self.date_folder,
            'pdf_reports': self.project_root / 'EHC_Data_Pdf' / self.date_folder,
            'logs': self.project_root / 'EHC_Logs' / self.date_folder,
            
            # Base data directories (without date folder)
            'csv_data_base': self.project_root / 'EHC_Data',
            'excel_merge_base': self.project_root / 'EHC_Data_Merge',
            'pdf_reports_base': self.project_root / 'EHC_Data_Pdf',
            'logs_base': self.project_root / 'EHC_Logs',
        }
    
    def ensure_directories(self):
        """Create all necessary directories"""
        paths = self.get_paths()
        
        created_dirs = []
        
        # Create data directories
        for key in ['csv_data', 'excel_merge', 'pdf_reports', 'logs']:
            path = paths[key]
            if not path.exists():
                try:
                    path.mkdir(parents=True, exist_ok=True)
                    created_dirs.append(str(path))
                except Exception as e:
                    print(f"❌ Failed to create {key}: {e}")
        
        if created_dirs:
            print(f"📁 Created directories:")
            for dir_path in created_dirs:
                print(f"   • {dir_path}")
        
        return paths
    
    def get_script_path(self, script_name: str) -> Path:
        """Get the full path to a script file"""
        paths = self.get_paths()
        
        # Common script locations
        script_locations = {
            # VBS scripts
            'vbs_phase1_login.py': paths['vbs_dir'] / 'vbs_phase1_login.py',
            'vbs_phase2_navigation_fixed.py': paths['vbs_dir'] / 'vbs_phase2_navigation_fixed.py', 
            'vbs_phase3_upload_complete.py': paths['vbs_dir'] / 'vbs_phase3_upload_complete.py',
            'vbs_phase4_report_fixed.py': paths['vbs_dir'] / 'vbs_phase4_report_fixed.py',
            
            # WiFi scripts
            'csv_downloader_resilient.py': paths['wifi_dir'] / 'csv_downloader_resilient.py',
            
            # Email scripts
            'outlook_automation.py': paths['email_dir'] / 'outlook_automation.py',
            'email_delivery.py': paths['email_dir'] / 'email_delivery.py',
            
            # Excel scripts
            'excel_generator.py': paths['excel_dir'] / 'excel_generator.py',
            
            # Utils scripts
            'close_vbs.py': paths['utils_dir'] / 'close_vbs.py',
        }
        
        if script_name in script_locations:
            return script_locations[script_name]
        
        # If not found in common locations, search project
        for script_dir in [paths['vbs_dir'], paths['wifi_dir'], paths['email_dir'], paths['excel_dir'], paths['utils_dir']]:
            script_path = script_dir / script_name
            if script_path.exists():
                return script_path
        
        # Fallback: return relative to project root
        return paths['project_root'] / script_name
    
    def cd_to_project_root(self):
        """Change to project root directory"""
        try:
            os.chdir(self.project_root)
            print(f"📂 Changed to project root: {self.project_root}")
            return True
        except Exception as e:
            print(f"❌ Failed to change to project root: {e}")
            return False

# Global instance for easy import
path_manager = UniversalPathManager()

def get_paths():
    """Quick access to paths"""
    return path_manager.get_paths()

def ensure_directories():
    """Quick access to directory creation"""
    return path_manager.ensure_directories()

def get_script_path(script_name: str):
    """Quick access to script paths"""
    return path_manager.get_script_path(script_name)

def cd_to_project_root():
    """Quick access to directory change"""
    return path_manager.cd_to_project_root()

if __name__ == "__main__":
    # Test the path manager
    print("\n🧪 Testing Universal Path Manager:")
    
    paths = ensure_directories()
    
    print(f"\n📋 Available paths:")
    for key, path in paths.items():
        exists = "✅" if path.exists() else "❌"
        print(f"   {key}: {exists} {path}")
    
    print(f"\n🔍 Script locations:")
    test_scripts = ['vbs_phase1_login.py', 'csv_downloader_resilient.py', 'excel_generator.py']
    for script in test_scripts:
        script_path = get_script_path(script)
        exists = "✅" if script_path.exists() else "❌"
        print(f"   {script}: {exists} {script_path}")

