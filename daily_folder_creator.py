#!/usr/bin/env python3
"""
Daily Folder Creator - MoonFlower Automation System
Creates daily folders for EHC_Data, EHC_Data_Merge, and EHC_Data_Pdf at 12:00 AM
Ensures proper directory structure for automation workflow
"""

import os
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path
import time

class DailyFolderCreator:
    """Creates and manages daily folder structure"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        self.base_directories = [
            "EHC_Data",
            "EHC_Data_Merge", 
            "EHC_Data_Pdf",
            "EHC_Logs"
        ]
        
    def _setup_logging(self):
        """Setup logging for folder creation"""
        logger = logging.getLogger("DailyFolderCreator")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def get_date_folder(self, date_obj=None):
        """Get date folder name in ddMMM format (e.g., 22jul)"""
        if date_obj is None:
            date_obj = datetime.now()
        return date_obj.strftime("%d%b").lower()
    
    def create_daily_folders(self, target_date=None):
        """Create daily folders for all base directories"""
        try:
            if target_date is None:
                target_date = datetime.now()
            
            date_folder = self.get_date_folder(target_date)
            created_folders = []
            
            self.logger.info(f"🗂️ Creating daily folders for {target_date.strftime('%Y-%m-%d')} ({date_folder})")
            
            for base_dir in self.base_directories:
                folder_path = Path(base_dir) / date_folder
                
                try:
                    folder_path.mkdir(parents=True, exist_ok=True)
                    created_folders.append(str(folder_path))
                    self.logger.info(f"✅ Created/ensured: {folder_path}")
                    
                    # Create debug_images subfolder for EHC_Logs
                    if base_dir == "EHC_Logs":
                        debug_path = folder_path / "debug_images"
                        debug_path.mkdir(exist_ok=True)
                        self.logger.info(f"✅ Created/ensured: {debug_path}")
                        
                except Exception as e:
                    self.logger.error(f"❌ Failed to create {folder_path}: {e}")
            
            self.logger.info(f"🎉 Daily folder creation completed for {date_folder}")
            return {"success": True, "date_folder": date_folder, "created_folders": created_folders}
            
        except Exception as e:
            self.logger.error(f"❌ Daily folder creation failed: {e}")
            return {"success": False, "error": str(e)}
    
    def create_future_folders(self, days_ahead=7):
        """Create folders for upcoming days"""
        try:
            self.logger.info(f"📅 Creating folders for next {days_ahead} days")
            
            for i in range(days_ahead):
                future_date = datetime.now() + timedelta(days=i)
                result = self.create_daily_folders(future_date)
                if not result["success"]:
                    self.logger.warning(f"⚠️ Failed to create folders for {future_date.strftime('%Y-%m-%d')}")
            
            self.logger.info(f"✅ Future folder creation completed")
            return {"success": True}
            
        except Exception as e:
            self.logger.error(f"❌ Future folder creation failed: {e}")
            return {"success": False, "error": str(e)}
    
    def cleanup_old_folders(self, days_to_keep=30):
        """Clean up folders older than specified days"""
        try:
            self.logger.info(f"🧹 Cleaning up folders older than {days_to_keep} days")
            
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            removed_folders = []
            
            for base_dir in self.base_directories:
                base_path = Path(base_dir)
                if not base_path.exists():
                    continue
                
                for folder in base_path.iterdir():
                    if folder.is_dir():
                        try:
                            # Parse folder name (ddMMM format)
                            folder_date = datetime.strptime(f"{folder.name}2025", "%d%b%Y")
                            
                            if folder_date < cutoff_date:
                                # Remove old folder
                                import shutil
                                shutil.rmtree(folder)
                                removed_folders.append(str(folder))
                                self.logger.info(f"🗑️ Removed old folder: {folder}")
                                
                        except ValueError:
                            # Skip folders that don't match date format
                            continue
                        except Exception as e:
                            self.logger.warning(f"⚠️ Failed to remove {folder}: {e}")
            
            self.logger.info(f"✅ Cleanup completed - removed {len(removed_folders)} folders")
            return {"success": True, "removed_folders": removed_folders}
            
        except Exception as e:
            self.logger.error(f"❌ Cleanup failed: {e}")
            return {"success": False, "error": str(e)}
    
    def validate_folder_structure(self):
        """Validate that today's folder structure exists"""
        try:
            today_folder = self.get_date_folder()
            missing_folders = []
            
            for base_dir in self.base_directories:
                folder_path = Path(base_dir) / today_folder
                if not folder_path.exists():
                    missing_folders.append(str(folder_path))
            
            if missing_folders:
                self.logger.warning(f"⚠️ Missing folders: {missing_folders}")
                return {"valid": False, "missing_folders": missing_folders}
            else:
                self.logger.info("✅ Folder structure validation passed")
                return {"valid": True, "missing_folders": []}
                
        except Exception as e:
            self.logger.error(f"❌ Folder validation failed: {e}")
            return {"valid": False, "error": str(e)}

def main():
    """Main function for command line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Daily Folder Creator for MoonFlower Automation")
    parser.add_argument('--create-today', action='store_true', help='Create folders for today')
    parser.add_argument('--create-future', type=int, default=7, help='Create folders for future days (default: 7)')
    parser.add_argument('--cleanup', type=int, default=30, help='Clean up folders older than N days (default: 30)')
    parser.add_argument('--validate', action='store_true', help='Validate folder structure')
    parser.add_argument('--all', action='store_true', help='Run all operations')
    
    args = parser.parse_args()
    
    creator = DailyFolderCreator()
    
    print("🗂️ MoonFlower Daily Folder Creator")
    print("=" * 50)
    
    if args.create_today or args.all:
        print("\n📅 Creating today's folders...")
        result = creator.create_daily_folders()
        if result["success"]:
            print(f"✅ Created folders for: {result['date_folder']}")
        else:
            print(f"❌ Failed: {result.get('error', 'Unknown error')}")
    
    if args.create_future or args.all:
        print(f"\n📅 Creating future folders ({args.create_future} days)...")
        result = creator.create_future_folders(args.create_future)
        if result["success"]:
            print("✅ Future folders created")
        else:
            print(f"❌ Failed: {result.get('error', 'Unknown error')}")
    
    if args.cleanup or args.all:
        print(f"\n🧹 Cleaning up old folders (>{args.cleanup} days)...")
        result = creator.cleanup_old_folders(args.cleanup)
        if result["success"]:
            print(f"✅ Cleaned up {len(result.get('removed_folders', []))} old folders")
        else:
            print(f"❌ Failed: {result.get('error', 'Unknown error')}")
    
    if args.validate or args.all:
        print("\n✅ Validating folder structure...")
        result = creator.validate_folder_structure()
        if result["valid"]:
            print("✅ Folder structure is valid")
        else:
            print(f"❌ Missing folders: {result.get('missing_folders', [])}")
    
    print("\n" + "=" * 50)
    print("Daily Folder Creator completed")

if __name__ == "__main__":
    main() 