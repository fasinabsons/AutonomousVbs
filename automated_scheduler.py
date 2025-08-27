#!/usr/bin/env python3
"""
Automated CSV Download Scheduler
Runs continuously and automatically downloads CSV files during slot times,
then merges them and sends notifications when complete.
"""

import sys
import time
import schedule
import logging
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Any, List
import json

# Add current directory to path for imports
sys.path.append(str(Path(__file__).parent))

from utils.dependency_manager import DependencyManager
from utils.config_loader import ConfigLoader
from enhanced_csv_cli import EnhancedCSVCLI

class AutomatedScheduler:
    """Automated scheduler for CSV downloads with time-based execution"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        self.config = ConfigLoader()
        self.is_running = False
        self.current_session = None
        self.download_completed = {"morning": False, "afternoon": False}
        self.merge_completed = False
        self.email_sent = False
        
        # Load time slots from configuration
        self.time_slots = self.config.get_time_slots()
        
        # Directory setup
        self.setup_directories()
        
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for the scheduler"""
        logger = logging.getLogger("AutomatedScheduler")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                date_folder = datetime.now().strftime('%d%b').lower()
                log_dir = Path("EHC_Logs") / date_folder
                log_dir.mkdir(parents=True, exist_ok=True)
                
                log_file = log_dir / f"automated_scheduler_{datetime.now().strftime('%Y%m%d')}.log"
                file_handler = logging.FileHandler(log_file)
                file_formatter = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
                )
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
                
            except Exception as e:
                logger.warning(f"Could not setup file logging: {e}")
        
        return logger
    
    def setup_directories(self):
        """Setup all required directories using configuration"""
        date_folder = datetime.now().strftime('%d%b').lower()
        
        # Create directories using configuration
        self.folder_paths = self.config.create_directories(date_folder)
        
        # Create debug images subdirectory
        debug_dir = Path(self.folder_paths['log_dir']) / "debug_images"
        debug_dir.mkdir(parents=True, exist_ok=True)
        self.logger.info(f"Created debug directory: {debug_dir}")
    
    def validate_environment(self) -> bool:
        """Validate environment and dependencies"""
        try:
            self.logger.info("=== VALIDATING ENVIRONMENT ===")
            
            # Check dependencies
            dependency_manager = DependencyManager()
            is_valid, missing_critical = dependency_manager.validate_critical_modules()
            
            if not is_valid:
                self.logger.error(f"Missing critical modules: {', '.join(missing_critical)}")
                self.logger.info("Attempting automatic dependency installation...")
                
                install_success = dependency_manager.install_missing_dependencies()
                if install_success:
                    self.logger.info("Dependencies installed successfully")
                    is_valid, missing_critical = dependency_manager.validate_critical_modules()
                    
                    if is_valid:
                        self.logger.info("All critical modules now available")
                        return True
                    else:
                        self.logger.error(f"Still missing modules: {', '.join(missing_critical)}")
                        return False
                else:
                    self.logger.error("Failed to install missing dependencies")
                    return False
            else:
                self.logger.info("All critical modules are available")
                return True
                
        except Exception as e:
            self.logger.error(f"Environment validation failed: {e}")
            return False
    
    def is_slot_time(self, session_type: str) -> bool:
        """Check if current time is within the slot time for given session"""
        current_time = datetime.now().strftime('%H:%M')
        slot_config = self.time_slots.get(session_type, {})
        
        start_time = slot_config.get('start_time', '00:00')
        end_time = slot_config.get('end_time', '23:59')
        
        return start_time <= current_time <= end_time
    
    def should_download_now(self, session_type: str) -> bool:
        """Check if we should download now based on scheduled times - EXACT TIMING"""
        current_time = datetime.now().strftime('%H:%M')
        slot_config = self.time_slots.get(session_type, {})
        download_times = slot_config.get('download_times', [])
        
        # Check if current time matches any download time (within 1 minute for exact timing)
        current_dt = datetime.strptime(current_time, '%H:%M')
        
        for download_time in download_times:
            download_dt = datetime.strptime(download_time, '%H:%M')
            time_diff = abs((current_dt - download_dt).total_seconds())
            
            if time_diff <= 60:  # Within 1 minute for exact timing
                return True
        
        return False
    
    def get_next_slot_time(self, session_type: str) -> str:
        """Get the next scheduled slot time for the session"""
        current_time = datetime.now().strftime('%H:%M')
        slot_config = self.time_slots.get(session_type, {})
        download_times = slot_config.get('download_times', [])
        
        current_dt = datetime.strptime(current_time, '%H:%M')
        
        for download_time in download_times:
            download_dt = datetime.strptime(download_time, '%H:%M')
            if download_dt > current_dt:
                return download_time
        
        return "No more slots today"
    
    def execute_csv_download(self, session_type: str) -> Dict[str, Any]:
        """Execute CSV download for specified session"""
        try:
            self.logger.info(f"=== STARTING {session_type.upper()} CSV DOWNLOAD ===")
            
            date_folder = datetime.now().strftime('%d%b').lower()
            output_dir = self.folder_paths['csv_dir']
            
            # Initialize enhanced CLI
            cli = EnhancedCSVCLI()
            cli.logger = cli.setup_enhanced_logging(date_folder, session_type)
            
            # Validate environment
            if not cli.validate_environment():
                self.logger.error("Environment validation failed")
                return {"success": False, "error": "Environment validation failed"}
            
            # Create mock args for the CLI
            class MockArgs:
                def __init__(self):
                    self.session = session_type
                    self.output = output_dir
                    self.date = date_folder
                    self.merge = False  # We'll merge separately
                    self.vbs = False
                    self.email = False
                    self.silent = True
            
            args = MockArgs()
            
            # Execute workflow
            result = cli.execute_workflow(args)
            
            if result.get("success", False):
                self.logger.info(f"{session_type.title()} download completed successfully")
                self.download_completed[session_type] = True
                return result
            else:
                self.logger.error(f"{session_type.title()} download failed: {result.get('errors', [])}")
                return result
                
        except Exception as e:
            self.logger.error(f"CSV download execution failed: {e}")
            return {"success": False, "error": str(e)}
    
    def execute_merge_and_notify(self) -> Dict[str, Any]:
        """Execute Excel merge and email notification after all downloads complete"""
        try:
            self.logger.info("=== STARTING MERGE AND NOTIFICATION ===")
            
            date_folder = datetime.now().strftime('%d%b').lower()
            csv_dir = self.folder_paths['csv_dir']
            excel_dir = self.folder_paths['excel_dir']
            
            # Initialize enhanced CLI
            cli = EnhancedCSVCLI()
            cli.logger = cli.setup_enhanced_logging(date_folder, "complete")
            
            # Create mock args for complete workflow
            class MockArgs:
                def __init__(self):
                    self.session = "complete"
                    self.output = csv_dir
                    self.date = date_folder
                    self.merge = True
                    self.vbs = True
                    self.email = True
                    self.silent = True
            
            args = MockArgs()
            
            # Execute merge and notification
            result = cli.execute_workflow(args)
            
            if result.get("success", False):
                self.logger.info("Merge and notification completed successfully")
                self.merge_completed = True
                self.email_sent = result.get("email_sent", False)
                return result
            else:
                self.logger.error(f"Merge and notification failed: {result.get('errors', [])}")
                return result
                
        except Exception as e:
            self.logger.error(f"Merge and notification execution failed: {e}")
            return {"success": False, "error": str(e)}
    
    def check_and_execute_downloads(self):
        """Check if it's time to download and execute if needed - ACCURATE TIMING ONLY"""
        try:
            current_time = datetime.now().strftime('%H:%M')
            current_hour = datetime.now().hour
            current_minute = datetime.now().minute
            
            # MORNING SESSION - Only download at exact scheduled times
            if not self.download_completed["morning"]:
                if self.is_slot_time("morning") and self.should_download_now("morning"):
                    self.logger.info(f"[{current_time}] ⏰ SCHEDULED MORNING DOWNLOAD STARTING")
                    result = self.execute_csv_download("morning")
                    
                    if not result.get("success", False):
                        self.logger.warning("Morning download failed, will retry at next scheduled time")
                elif current_hour >= 13 and not self.download_completed["morning"]:
                    # Only download if we completely missed morning session
                    self.logger.info(f"[{current_time}] ⚠️ MISSED ALL MORNING SLOTS - Emergency download")
                    result = self.execute_csv_download("morning")
            
            # AFTERNOON SESSION - Only download at exact scheduled times  
            if not self.download_completed["afternoon"]:
                if self.is_slot_time("afternoon") and self.should_download_now("afternoon"):
                    self.logger.info(f"[{current_time}] ⏰ SCHEDULED AFTERNOON DOWNLOAD STARTING")
                    result = self.execute_csv_download("afternoon")
                    
                    if not result.get("success", False):
                        self.logger.warning("Afternoon download failed, will retry at next scheduled time")
                elif current_hour >= 18 and not self.download_completed["afternoon"]:
                    # Only download if we completely missed afternoon session
                    self.logger.info(f"[{current_time}] ⚠️ MISSED ALL AFTERNOON SLOTS - Emergency download")
                    result = self.execute_csv_download("afternoon")
            
            # MERGE AND EMAIL - Only after downloads are complete
            if (self.download_completed["morning"] and 
                self.download_completed["afternoon"] and 
                not self.merge_completed):
                
                self.logger.info("📊 Both sessions completed - Starting merge and email notification")
                result = self.execute_merge_and_notify()
                
                if result.get("success", False):
                    self.logger.info("✅ ALL OPERATIONS COMPLETED SUCCESSFULLY!")
                    email_config = self.config.get_email_config()
                    self.logger.info(f"📧 Email sent to {email_config['recipient']}")
                    self.logger.info("🔄 System will continue monitoring for next day...")
            
            # End of day merge (if only one session completed)
            elif current_hour >= 19 and not self.merge_completed:
                if self.download_completed["morning"] or self.download_completed["afternoon"]:
                    self.logger.info("🌙 End of day - Merging available downloads")
                    result = self.execute_merge_and_notify()
                    
                    if result.get("success", False):
                        self.logger.info("✅ END OF DAY OPERATIONS COMPLETED!")
                        email_config = self.config.get_email_config()
                        self.logger.info(f"📧 Email sent to {email_config['recipient']}")
                        self.logger.info("🔄 System will continue monitoring for next day...")
            
            return False  # Always continue monitoring
            
        except Exception as e:
            self.logger.error(f"Check and execute downloads failed: {e}")
            return False
    
    def reset_daily_status(self):
        """Reset daily status for new day"""
        self.download_completed = {"morning": False, "afternoon": False}
        self.merge_completed = False
        self.email_sent = False
        self.setup_directories()  # Create new day directories
        self.logger.info("Daily status reset for new day")
    
    def run_scheduler(self):
        """Run the automated scheduler"""
        try:
            current_time = datetime.now().strftime('%H:%M')
            current_hour = datetime.now().hour
            
            self.logger.info("=== AUTOMATED CSV SCHEDULER STARTING ===")
            self.logger.info(f"Current time: {current_time}")
            self.logger.info("System will automatically download CSV files during slot times")
            morning_times = ', '.join(self.time_slots['morning']['download_times'])
            afternoon_times = ', '.join(self.time_slots['afternoon']['download_times'])
            self.logger.info(f"Morning slots: {morning_times}")
            self.logger.info(f"Afternoon slots: {afternoon_times}")
            
            # Show what will happen based on current time
            if current_hour < 8:
                self.logger.info("⏰ Waiting for first morning slot at 08:30")
            elif current_hour < 13:
                self.logger.info("🌅 In morning session - will download at next slot or immediately if missed")
            elif current_hour < 17:
                self.logger.info("🌇 In afternoon session - will download at next slot or immediately if missed")
            else:
                self.logger.info("🌙 After hours - will download immediately if not done today")
            
            self.logger.info("Press Ctrl+C to stop the scheduler")
            
            # Validate environment first
            if not self.validate_environment():
                self.logger.error("Environment validation failed, cannot start scheduler")
                return False
            
            self.is_running = True
            last_date = datetime.now().date()
            
            # Schedule daily reset at midnight
            schedule.every().day.at("00:01").do(self.reset_daily_status)
            
            while self.is_running:
                try:
                    # Check if date changed (new day)
                    current_date = datetime.now().date()
                    if current_date != last_date:
                        self.reset_daily_status()
                        last_date = current_date
                        self.logger.info(f"New day started: {current_date}")
                    
                    # Run scheduled jobs
                    schedule.run_pending()
                    
                    # Check and execute downloads
                    self.check_and_execute_downloads()
                    
                    # Show status every 30 minutes
                    current_minute = datetime.now().minute
                    if current_minute % 30 == 0:
                        current_time = datetime.now().strftime('%H:%M')
                        morning_status = "✅ DONE" if self.download_completed["morning"] else "⏳ PENDING"
                        afternoon_status = "✅ DONE" if self.download_completed["afternoon"] else "⏳ PENDING"
                        merge_status = "✅ DONE" if self.merge_completed else "⏳ PENDING"
                        
                        self.logger.info(f"[{current_time}] Status - Morning: {morning_status}, Afternoon: {afternoon_status}, Merge: {merge_status}")
                    
                    # Sleep for configured interval before next check
                    check_interval = self.config.get('CHECK_INTERVAL', 1) * 60  # Convert minutes to seconds
                    time.sleep(check_interval)
                    
                except KeyboardInterrupt:
                    self.logger.info("Scheduler stopped by user (Ctrl+C)")
                    self.is_running = False
                    break
                    
                except Exception as e:
                    self.logger.error(f"Scheduler error: {e}")
                    self.logger.info("Continuing scheduler operation...")
                    time.sleep(60)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Scheduler failed to start: {e}")
            return False
    
    def run_once(self):
        """Run once immediately (for testing or manual execution)"""
        try:
            self.logger.info("=== RUNNING IMMEDIATE CSV DOWNLOAD ===")
            
            # Validate environment
            if not self.validate_environment():
                self.logger.error("Environment validation failed")
                return False
            
            # Run both sessions immediately
            morning_result = self.execute_csv_download("morning")
            afternoon_result = self.execute_csv_download("afternoon")
            
            # Mark as completed if successful
            if morning_result.get("success", False):
                self.download_completed["morning"] = True
            if afternoon_result.get("success", False):
                self.download_completed["afternoon"] = True
            
            # Run merge and notification
            if self.download_completed["morning"] or self.download_completed["afternoon"]:
                merge_result = self.execute_merge_and_notify()
                
                if merge_result.get("success", False):
                    self.logger.info("Immediate execution completed successfully!")
                    return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"Immediate execution failed: {e}")
            return False

def main():
    """Main entry point"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Automated CSV Download Scheduler")
    parser.add_argument('--mode', choices=['schedule', 'once'], default='schedule',
                       help='Run mode: schedule (continuous) or once (immediate)')
    parser.add_argument('--test', action='store_true',
                       help='Test mode - validate environment only')
    
    args = parser.parse_args()
    
    scheduler = AutomatedScheduler()
    
    if args.test:
        print("=== TESTING ENVIRONMENT ===")
        if scheduler.validate_environment():
            print("✅ Environment validation passed")
            return 0
        else:
            print("❌ Environment validation failed")
            return 1
    
    if args.mode == 'once':
        print("=== RUNNING IMMEDIATE DOWNLOAD ===")
        if scheduler.run_once():
            print("✅ Immediate execution completed successfully")
            return 0
        else:
            print("❌ Immediate execution failed")
            return 1
    
    else:  # schedule mode
        print("=== STARTING AUTOMATED SCHEDULER ===")
        print("The system will run continuously and download CSV files automatically")
        print("Press Ctrl+C to stop")
        
        if scheduler.run_scheduler():
            print("✅ Scheduler completed successfully")
            return 0
        else:
            print("❌ Scheduler failed")
            return 1

if __name__ == "__main__":
    sys.exit(main())