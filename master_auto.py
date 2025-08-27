#!/usr/bin/env python3
"""
🌙 MoonFlower Master Automation - Robust EXE Version
Simple, reliable 365-day automation with minimal UI configuration
"""

import os
import sys
import json
import time
import logging
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from pathlib import Path
import shutil

class MasterAutoConfig:
    """Simple configuration manager"""
    
    def __init__(self):
        self.config_file = "master_auto_config.json"
        self.default_config = {
            # Essential Configuration Only
            "email_settings": {
                "gm_email": "ramon.logan@absons.ae",
                "notification_email": "faseenm@gmail.com"
            },
            "locations": {
                "vbs_software_path": "",
                "chrome_download_path": "",
                "data_directory": ""
            },
            "timing": {
                "email_time": "08:30",
                "csv_slot1": "09:30", 
                "csv_slot2": "12:30",
                "vbs_upload": "12:40",
                "vbs_report": "17:01"
            }
        }
        self.load_config()
    
    def load_config(self):
        """Load configuration from file"""
        try:
            if Path(self.config_file).exists():
                with open(self.config_file, 'r') as f:
                    self.config = json.load(f)
                # Ensure all default keys exist
                self._merge_defaults()
            else:
                self.config = self.default_config.copy()
                self.save_config()
        except Exception as e:
            print(f"Config load error: {e}, using defaults")
            self.config = self.default_config.copy()
    
    def _merge_defaults(self):
        """Merge default values for missing keys"""
        for section, values in self.default_config.items():
            if section not in self.config:
                self.config[section] = values.copy()
            else:
                for key, value in values.items():
                    if key not in self.config[section]:
                        self.config[section][key] = value
    
    def save_config(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
            return True
        except Exception as e:
            print(f"Config save error: {e}")
            return False

class ScriptExecutor:
    """Robust script execution with error handling"""
    
    def __init__(self, logger):
        self.logger = logger
        self.setup_paths()
    
    def setup_paths(self):
        """Setup execution paths for embedded environment"""
        if getattr(sys, 'frozen', False):
            # Running as EXE
            self.exe_dir = Path(sys.executable).parent
            self.temp_dir = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else self.exe_dir
        else:
            # Running as script
            self.exe_dir = Path(__file__).parent
            self.temp_dir = self.exe_dir
    
    def execute_script(self, script_path, script_name, timeout=600):
        """Execute a Python script with robust error handling"""
        try:
            self.logger.info(f"🚀 Starting {script_name}: {script_path}")
            
            # Find script location (embedded or local)
            script_locations = [
                self.exe_dir / script_path,  # Copied to working dir
                self.temp_dir / script_path,  # Embedded in EXE
                Path(script_path)  # Direct path
            ]
            
            full_script_path = None
            for location in script_locations:
                if location.exists():
                    full_script_path = str(location)
                    break
            
            if not full_script_path:
                self.logger.error(f"❌ Script not found: {script_path}")
                return False
            
            # Execute with proper environment
            env = os.environ.copy()
            env['PYTHONPATH'] = str(self.exe_dir)
            
            result = subprocess.run([
                sys.executable, full_script_path
            ], 
            capture_output=True, 
            text=True, 
            timeout=timeout,
            cwd=str(self.exe_dir),
            env=env
            )
            
            if result.returncode == 0:
                self.logger.info(f"✅ {script_name} completed successfully")
                return True
            else:
                self.logger.error(f"❌ {script_name} failed (exit code: {result.returncode})")
                if result.stderr:
                    self.logger.error(f"Error: {result.stderr[:500]}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error(f"⏰ {script_name} timed out after {timeout} seconds")
            return False
        except Exception as e:
            self.logger.error(f"❌ {script_name} execution error: {e}")
            return False

class AutomationScheduler:
    """Simple, reliable scheduling engine"""
    
    def __init__(self, config, executor, logger):
        self.config = config
        self.executor = executor
        self.logger = logger
        self.running = False
        self.daily_tasks = {}
    
    def start(self):
        """Start the automation scheduler"""
        self.running = True
        self.logger.info("🕐 Automation scheduler started")
        
        # Start main loop in thread
        scheduler_thread = threading.Thread(target=self._scheduler_loop, daemon=True)
        scheduler_thread.start()
    
    def stop(self):
        """Stop the automation scheduler"""
        self.running = False
        self.logger.info("⏹️ Automation scheduler stopped")
    
    def _scheduler_loop(self):
        """Main scheduling loop"""
        while self.running:
            try:
                current_time = datetime.now().strftime("%H:%M")
                timing = self.config.config["timing"]
                
                # Check if tasks need to be reset (new day)
                self._check_new_day()
                
                # Morning Email
                if (current_time == timing["email_time"] and 
                    not self._task_done_today("email")):
                    self._execute_email()
                
                # CSV Slot 1
                elif (current_time == timing["csv_slot1"] and 
                      not self._task_done_today("csv_slot1")):
                    self._execute_csv_slot("csv_slot1")
                
                # CSV Slot 2
                elif (current_time == timing["csv_slot2"] and 
                      not self._task_done_today("csv_slot2")):
                    self._execute_csv_slot("csv_slot2")
                    # Auto-run Excel after slot 2
                    time.sleep(30)
                    self._execute_excel()
                
                # VBS Upload
                elif (current_time == timing["vbs_upload"] and 
                      not self._task_done_today("vbs_upload")):
                    self._execute_vbs_upload()
                
                # VBS Report
                elif (current_time == timing["vbs_report"] and 
                      not self._task_done_today("vbs_report")):
                    self._execute_vbs_report()
                
                # Sleep for 30 seconds
                time.sleep(30)
                
            except Exception as e:
                self.logger.error(f"❌ Scheduler error: {e}")
                time.sleep(60)
    
    def _check_new_day(self):
        """Reset daily tasks at midnight"""
        today = datetime.now().strftime("%Y-%m-%d")
        if not hasattr(self, '_current_day') or self._current_day != today:
            self._current_day = today
            self.daily_tasks = {}
            self.logger.info(f"🌅 New day started: {today}")
    
    def _task_done_today(self, task_name):
        """Check if task was completed today"""
        return self.daily_tasks.get(task_name, False)
    
    def _mark_task_done(self, task_name):
        """Mark task as completed"""
        self.daily_tasks[task_name] = True
        self.logger.info(f"✅ Task completed: {task_name}")
    
    def _execute_email(self):
        """Execute morning email"""
        self.logger.info("📧 Sending morning email to GM")
        if self.executor.execute_script("email/outlook_automation.py", "Morning Email"):
            self._mark_task_done("email")
        else:
            self._send_error_notification("Email Failed", "Morning email to GM failed")
    
    def _execute_csv_slot(self, slot_name):
        """Execute CSV download slot"""
        self.logger.info(f"📥 Starting {slot_name}")
        if self.executor.execute_script("wifi/csv_downloader_resilient.py", f"CSV {slot_name}"):
            self._mark_task_done(slot_name)
        else:
            self._send_error_notification(f"{slot_name} Failed", f"CSV download {slot_name} failed")
    
    def _execute_excel(self):
        """Execute Excel generation"""
        self.logger.info("📊 Generating Excel file")
        if self.executor.execute_script("excel/excel_generator.py", "Excel Generation"):
            self._mark_task_done("excel")
        else:
            self._send_error_notification("Excel Failed", "Excel generation failed")
    
    def _execute_vbs_upload(self):
        """Execute VBS upload workflow"""
        self.logger.info("⬆️ Starting VBS upload")
        scripts = [
            ("vbs/vbs_phase1_login.py", "VBS Login"),
            ("vbs/vbs_phase2_navigation_fixed.py", "VBS Navigation"), 
            ("vbs/vbs_phase3_upload_complete.py", "VBS Upload")
        ]
        
        success = True
        for script, name in scripts:
            if not self.executor.execute_script(script, name, timeout=1200):
                success = False
                break
            time.sleep(10)
        
        if success:
            self._mark_task_done("vbs_upload")
        else:
            self._send_error_notification("VBS Upload Failed", "VBS upload workflow failed")
    
    def _execute_vbs_report(self):
        """Execute VBS report generation"""
        self.logger.info("📊 Generating VBS report")
        scripts = [
            ("vbs/vbs_phase1_login.py", "VBS Login"),
            ("vbs/vbs_phase4_report_fixed.py", "VBS Report")
        ]
        
        success = True
        for script, name in scripts:
            if not self.executor.execute_script(script, name, timeout=600):
                success = False
                break
            time.sleep(10)
        
        if success:
            self._mark_task_done("vbs_report")
        else:
            self._send_error_notification("VBS Report Failed", "VBS report generation failed")
    
    def _send_error_notification(self, subject, message):
        """Send error notification email"""
        try:
            self.executor.execute_script("email/email_delivery.py", "Error Notification")
            self.logger.warning(f"🚨 {subject}: {message}")
        except Exception as e:
            self.logger.error(f"❌ Failed to send error notification: {e}")

class SimpleConfigUI:
    """Simple UI for essential configuration only"""
    
    def __init__(self, config):
        self.config = config
        self.root = None
        self.vars = {}
    
    def show(self):
        """Show configuration window"""
        self.root = tk.Tk()
        self.root.title("🌙 MoonFlower Configuration")
        self.root.geometry("600x500")
        self.root.resizable(False, False)
        
        # Main notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create tabs
        self._create_email_tab(notebook)
        self._create_locations_tab(notebook) 
        self._create_timing_tab(notebook)
        
        # Save button
        save_frame = tk.Frame(self.root)
        save_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Button(save_frame, text="💾 Save Configuration", 
                 command=self._save_config, bg="#27ae60", fg="white",
                 font=("Arial", 12, "bold"), height=2).pack(fill="x")
        
        self.root.mainloop()
    
    def _create_email_tab(self, notebook):
        """Create email configuration tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📧 Email Settings")
        
        # GM Email
        tk.Label(frame, text="General Manager Email:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.vars["gm_email"] = tk.StringVar(value=self.config.config["email_settings"]["gm_email"])
        tk.Entry(frame, textvariable=self.vars["gm_email"], font=("Arial", 10), width=50).pack(padx=10, pady=5)
        
        # Notification Email
        tk.Label(frame, text="Error Notification Email:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        self.vars["notification_email"] = tk.StringVar(value=self.config.config["email_settings"]["notification_email"])
        tk.Entry(frame, textvariable=self.vars["notification_email"], font=("Arial", 10), width=50).pack(padx=10, pady=5)
    
    def _create_locations_tab(self, notebook):
        """Create locations configuration tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📁 Locations")
        
        # VBS Software Path
        tk.Label(frame, text="VBS Software Path:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        vbs_frame = tk.Frame(frame)
        vbs_frame.pack(fill="x", padx=10, pady=5)
        self.vars["vbs_software_path"] = tk.StringVar(value=self.config.config["locations"]["vbs_software_path"])
        tk.Entry(vbs_frame, textvariable=self.vars["vbs_software_path"], font=("Arial", 9), width=45).pack(side="left", fill="x", expand=True)
        tk.Button(vbs_frame, text="Browse", command=lambda: self._browse_file("vbs_software_path")).pack(side="right", padx=(5,0))
        
        # Chrome Download Path
        tk.Label(frame, text="Chrome Download Directory:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        chrome_frame = tk.Frame(frame)
        chrome_frame.pack(fill="x", padx=10, pady=5)
        self.vars["chrome_download_path"] = tk.StringVar(value=self.config.config["locations"]["chrome_download_path"])
        tk.Entry(chrome_frame, textvariable=self.vars["chrome_download_path"], font=("Arial", 9), width=45).pack(side="left", fill="x", expand=True)
        tk.Button(chrome_frame, text="Browse", command=lambda: self._browse_folder("chrome_download_path")).pack(side="right", padx=(5,0))
        
        # Data Directory
        tk.Label(frame, text="Data Directory (EHC folders):", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        data_frame = tk.Frame(frame)
        data_frame.pack(fill="x", padx=10, pady=5)
        self.vars["data_directory"] = tk.StringVar(value=self.config.config["locations"]["data_directory"])
        tk.Entry(data_frame, textvariable=self.vars["data_directory"], font=("Arial", 9), width=45).pack(side="left", fill="x", expand=True)
        tk.Button(data_frame, text="Browse", command=lambda: self._browse_folder("data_directory")).pack(side="right", padx=(5,0))
    
    def _create_timing_tab(self, notebook):
        """Create timing configuration tab"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="⏰ Timing")
        
        timings = [
            ("Morning Email Time", "email_time"),
            ("CSV Slot 1 Time", "csv_slot1"),
            ("CSV Slot 2 Time", "csv_slot2"), 
            ("VBS Upload Time", "vbs_upload"),
            ("VBS Report Time", "vbs_report")
        ]
        
        for label, key in timings:
            tk.Label(frame, text=f"{label}:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
            self.vars[key] = tk.StringVar(value=self.config.config["timing"][key])
            entry = tk.Entry(frame, textvariable=self.vars[key], font=("Arial", 10), width=10)
            entry.pack(anchor="w", padx=10, pady=2)
    
    def _browse_file(self, var_name):
        """Browse for file"""
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            self.vars[var_name].set(filename)
    
    def _browse_folder(self, var_name):
        """Browse for folder"""
        folder = filedialog.askdirectory(title="Select Directory")
        if folder:
            self.vars[var_name].set(folder)
    
    def _save_config(self):
        """Save configuration and close window"""
        try:
            # Update config object
            self.config.config["email_settings"]["gm_email"] = self.vars["gm_email"].get()
            self.config.config["email_settings"]["notification_email"] = self.vars["notification_email"].get()
            self.config.config["locations"]["vbs_software_path"] = self.vars["vbs_software_path"].get()
            self.config.config["locations"]["chrome_download_path"] = self.vars["chrome_download_path"].get()
            self.config.config["locations"]["data_directory"] = self.vars["data_directory"].get()
            self.config.config["timing"]["email_time"] = self.vars["email_time"].get()
            self.config.config["timing"]["csv_slot1"] = self.vars["csv_slot1"].get()
            self.config.config["timing"]["csv_slot2"] = self.vars["csv_slot2"].get()
            self.config.config["timing"]["vbs_upload"] = self.vars["vbs_upload"].get()
            self.config.config["timing"]["vbs_report"] = self.vars["vbs_report"].get()
            
            if self.config.save_config():
                messagebox.showinfo("Success", "Configuration saved successfully!")
                self.root.destroy()
            else:
                messagebox.showerror("Error", "Failed to save configuration!")
                
        except Exception as e:
            messagebox.showerror("Error", f"Configuration save error: {e}")

class MasterAuto:
    """Main automation controller"""
    
    def __init__(self):
        self.setup_logging()
        self.setup_environment()
        self.config = MasterAutoConfig()
        self.executor = ScriptExecutor(self.logger)
        self.scheduler = AutomationScheduler(self.config, self.executor, self.logger)
        self.running = False
    
    def setup_logging(self):
        """Setup logging system"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_dir / f"master_auto_{datetime.now().strftime('%Y%m%d')}.log", encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger("MasterAuto")
    
    def setup_environment(self):
        """Setup execution environment"""
        # Create necessary directories
        dirs = ["EHC_Data", "EHC_Data_Merge", "EHC_Data_Pdf", "EHC_Logs", "logs"]
        for dir_name in dirs:
            Path(dir_name).mkdir(exist_ok=True)
        
        # Copy embedded files if running from EXE
        self._copy_embedded_files()
    
    def _copy_embedded_files(self):
        """Copy embedded Python files from EXE to working directory"""
        if not getattr(sys, 'frozen', False):
            return  # Not running as EXE
        
        try:
            temp_dir = Path(sys._MEIPASS) if hasattr(sys, '_MEIPASS') else Path.cwd()
            exe_dir = Path(sys.executable).parent
            
            # Critical Python files to copy
            critical_files = [
                "daily_folder_creator.py",
                "email/outlook_automation.py",
                "email/email_delivery.py",
                "wifi/csv_downloader_resilient.py", 
                "excel/excel_generator.py",
                "vbs/vbs_phase1_login.py",
                "vbs/vbs_phase2_navigation_fixed.py",
                "vbs/vbs_phase3_upload_complete.py",
                "vbs/vbs_phase4_report_fixed.py"
            ]
            
            for file_path in critical_files:
                source = temp_dir / file_path
                target = exe_dir / file_path
                
                if source.exists():
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, target)
                    self.logger.info(f"✅ Copied {file_path}")
            
            # Copy Images directory
            images_source = temp_dir / "Images"
            images_target = exe_dir / "Images"
            if images_source.exists():
                if images_target.exists():
                    shutil.rmtree(images_target)
                shutil.copytree(images_source, images_target)
                self.logger.info("✅ Images directory copied")
                
        except Exception as e:
            self.logger.error(f"❌ File copying failed: {e}")
    
    def show_config(self):
        """Show configuration UI"""
        ui = SimpleConfigUI(self.config)
        ui.show()
    
    def start_automation(self):
        """Start the automation system"""
        if self.running:
            return False
        
        self.running = True
        self.logger.info("🚀 MasterAuto starting...")
        
        # Start folder creator
        self.executor.execute_script("daily_folder_creator.py", "Daily Folder Creator")
        
        # Start scheduler
        self.scheduler.start()
        
        self.logger.info("✅ MasterAuto automation started - 365-day operation active")
        return True
    
    def stop_automation(self):
        """Stop the automation system"""
        if not self.running:
            return False
        
        self.running = False
        self.scheduler.stop()
        self.logger.info("⏹️ MasterAuto automation stopped")
        return True

def main():
    """Main entry point"""
    try:
        app = MasterAuto()
        
        # Check command line arguments
        if len(sys.argv) > 1:
            if sys.argv[1] == "--config":
                app.show_config()
                return
            elif sys.argv[1] == "--start":
                app.start_automation()
                # Keep running
                try:
                    while app.running:
                        time.sleep(1)
                except KeyboardInterrupt:
                    app.stop_automation()
                return
        
        # Default: Show simple choice
        root = tk.Tk()
        root.title("🌙 MoonFlower Master Automation")
        root.geometry("400x200")
        root.resizable(False, False)
        
        # Center window
        root.eval('tk::PlaceWindow . center')
        
        # Title
        tk.Label(root, text="🌙 MoonFlower Automation", 
                font=("Arial", 16, "bold")).pack(pady=20)
        
        # Buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="⚙️ Configuration", 
                 command=app.show_config, bg="#3498db", fg="white",
                 font=("Arial", 12), width=15).pack(pady=5)
        
        def start_automation():
            if app.start_automation():
                messagebox.showinfo("Started", "Automation started! Running in background.")
                root.destroy()
        
        tk.Button(button_frame, text="🚀 Start Automation", 
                 command=start_automation, bg="#27ae60", fg="white",
                 font=("Arial", 12), width=15).pack(pady=5)
        
        root.mainloop()
        
    except Exception as e:
        print(f"Critical error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
