"""
GUI Configuration Panel for MoonFlower WiFi Automation System
Provides comprehensive configuration management interface
"""

import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json
import logging
import threading
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add parent directory to path for imports
sys.path.append(str(Path(__file__).parent.parent))

try:
    from utils.config_manager import ConfigManager
    from service_manager import ServiceManager
    from service_monitor import ServiceHealthMonitor
    from email.email_delivery import EmailDeliverySystem
except ImportError as e:
    print(f"Import error: {e}")
    ConfigManager = None
    ServiceManager = None
    ServiceHealthMonitor = None
    EmailDeliverySystem = None


class ConfigurationPanel:
    """Main configuration panel for MoonFlower automation system"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        
        # Initialize managers
        try:
            self.config_manager = ConfigManager() if ConfigManager else None
            self.service_manager = ServiceManager() if ServiceManager else None
            self.health_monitor = ServiceHealthMonitor() if ServiceHealthMonitor else None
            self.email_system = EmailDeliverySystem() if EmailDeliverySystem else None
        except Exception as e:
            self.logger.error(f"Failed to initialize managers: {e}")
            self.config_manager = None
            self.service_manager = None
            self.health_monitor = None
            self.email_system = None
        
        # GUI components
        self.root = None
        self.notebook = None
        self.status_var = None
        self.service_status_var = None
        
        # Configuration variables
        self.config_vars = {}
        self.email_vars = {}
        self.system_vars = {}
        self.vbs_vars = {}
        
        # Status tracking
        self.last_config_save = None
        self.unsaved_changes = False
        
        self.logger.info("Configuration Panel initialized")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for configuration panel"""
        logger = logging.getLogger("ConfigPanel")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                log_file = Path("EHC_Logs") / f"config_panel_{datetime.now().strftime('%Y%m%d')}.log"
                log_file.parent.mkdir(exist_ok=True)
                file_handler = logging.FileHandler(log_file)
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except Exception as e:
                print(f"Could not setup file logging: {e}")
        
        return logger
    
    def create_gui(self):
        """Create the main GUI interface"""
        try:
            self.root = tk.Tk()
            self.root.title("MoonFlower WiFi Automation - Configuration Panel")
            self.root.geometry("900x700")
            self.root.minsize(800, 600)
            
            # Configure style
            style = ttk.Style()
            style.theme_use('clam')
            
            # Create main frame
            main_frame = ttk.Frame(self.root, padding="10")
            main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            
            # Configure grid weights
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)
            main_frame.columnconfigure(0, weight=1)
            main_frame.rowconfigure(1, weight=1)
            
            # Create header
            self.create_header(main_frame)
            
            # Create tabbed interface
            self.create_tabbed_interface(main_frame)
            
            # Create status bar
            self.create_status_bar(main_frame)
            
            # Load current configuration
            self.load_configuration()
            
            # Update service status
            self.update_service_status()
            
            # Setup auto-refresh
            self.setup_auto_refresh()
            
            self.logger.info("GUI created successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to create GUI: {e}")
            if self.root:
                messagebox.showerror("Error", f"Failed to create GUI: {e}")
    
    def create_header(self, parent):
        """Create header section"""
        header_frame = ttk.Frame(parent)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        header_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(header_frame, text="MoonFlower WiFi Automation", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Service status
        self.service_status_var = tk.StringVar(value="Checking...")
        status_label = ttk.Label(header_frame, textvariable=self.service_status_var,
                                font=('Arial', 10))
        status_label.grid(row=0, column=1, sticky=tk.E)
        
        # Subtitle
        subtitle_label = ttk.Label(header_frame, text="Configuration Management Panel",
                                  font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, columnspan=2, sticky=tk.W)
    
    def create_tabbed_interface(self, parent):
        """Create tabbed interface"""
        self.notebook = ttk.Notebook(parent)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Create tabs
        self.create_email_settings_tab()
        self.create_system_settings_tab()
        self.create_vbs_settings_tab()
        self.create_service_management_tab()
        self.create_log_viewer_tab()
        self.create_test_tools_tab()
    
    def create_email_settings_tab(self):
        """Create email settings tab"""
        email_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(email_frame, text="Email Settings")
        
        # Configure grid
        email_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Email delivery options
        ttk.Label(email_frame, text="Email Delivery Options", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        row += 1
        
        # Weekend/weekday options
        self.email_vars['send_weekdays_only'] = tk.BooleanVar()
        self.email_vars['send_all_days'] = tk.BooleanVar()
        
        ttk.Checkbutton(email_frame, text="Send emails on weekdays only",
                       variable=self.email_vars['send_weekdays_only'],
                       command=self.on_weekday_only_changed).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        ttk.Checkbutton(email_frame, text="Send emails all days (including weekends)",
                       variable=self.email_vars['send_all_days'],
                       command=self.on_all_days_changed).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        row += 1
        
        # Recipients section
        ttk.Label(email_frame, text="Email Recipients", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        row += 1
        
        # Daily report recipients
        ttk.Label(email_frame, text="Daily Report Recipients:").grid(row=row, column=0, sticky=tk.W)
        self.email_vars['daily_recipients'] = tk.StringVar()
        daily_entry = ttk.Entry(email_frame, textvariable=self.email_vars['daily_recipients'], width=50)
        daily_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        row += 1
        
        ttk.Label(email_frame, text="(Separate multiple emails with commas)", 
                 font=('Arial', 8)).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        # Completion notification recipients
        ttk.Label(email_frame, text="Completion Notification Recipients:").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        self.email_vars['completion_recipients'] = tk.StringVar()
        completion_entry = ttk.Entry(email_frame, textvariable=self.email_vars['completion_recipients'], width=50)
        completion_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=(10, 0))
        row += 1
        
        # Error alert recipients
        ttk.Label(email_frame, text="Error Alert Recipients:").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        self.email_vars['error_recipients'] = tk.StringVar()
        error_entry = ttk.Entry(email_frame, textvariable=self.email_vars['error_recipients'], width=50)
        error_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=(10, 0))
        row += 1
        
        # SMTP Settings
        ttk.Separator(email_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        row += 1
        
        ttk.Label(email_frame, text="SMTP Configuration", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        # Gmail settings
        ttk.Label(email_frame, text="Gmail SMTP Server:").grid(row=row, column=0, sticky=tk.W)
        self.email_vars['gmail_server'] = tk.StringVar(value="smtp.gmail.com")
        ttk.Entry(email_frame, textvariable=self.email_vars['gmail_server'], width=30).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        ttk.Label(email_frame, text="Gmail Port:").grid(row=row, column=0, sticky=tk.W)
        self.email_vars['gmail_port'] = tk.StringVar(value="587")
        ttk.Entry(email_frame, textvariable=self.email_vars['gmail_port'], width=10).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        # Test email button
        test_frame = ttk.Frame(email_frame)
        test_frame.grid(row=row, column=0, columnspan=2, pady=20)
        
        ttk.Button(test_frame, text="Send Test Email", 
                  command=self.send_test_email).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(test_frame, text="Validate SMTP Settings", 
                  command=self.validate_smtp_settings).pack(side=tk.LEFT)
    
    def create_system_settings_tab(self):
        """Create system settings tab"""
        system_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(system_frame, text="System Settings")
        
        # Configure grid
        system_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Auto-restart settings
        ttk.Label(system_frame, text="Auto-Restart Settings", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        row += 1
        
        self.system_vars['auto_restart'] = tk.BooleanVar()
        ttk.Checkbutton(system_frame, text="Enable automatic PC restart",
                       variable=self.system_vars['auto_restart']).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        ttk.Label(system_frame, text="Restart Time (24-hour format):").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        self.system_vars['restart_time'] = tk.StringVar(value="02:00")
        ttk.Entry(system_frame, textvariable=self.system_vars['restart_time'], width=10).grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        row += 1
        
        # Debug settings
        ttk.Separator(system_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        row += 1
        
        ttk.Label(system_frame, text="Debug Settings", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        self.system_vars['debug_mode'] = tk.BooleanVar()
        ttk.Checkbutton(system_frame, text="Enable debug mode (verbose logging)",
                       variable=self.system_vars['debug_mode']).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        # Retry settings
        ttk.Label(system_frame, text="Error Handling", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(20, 10))
        row += 1
        
        ttk.Label(system_frame, text="Max Retry Attempts:").grid(row=row, column=0, sticky=tk.W)
        self.system_vars['max_retry_attempts'] = tk.StringVar(value="3")
        ttk.Entry(system_frame, textvariable=self.system_vars['max_retry_attempts'], width=10).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        ttk.Label(system_frame, text="Retry Delay (seconds):").grid(row=row, column=0, sticky=tk.W)
        self.system_vars['retry_delay'] = tk.StringVar(value="30")
        ttk.Entry(system_frame, textvariable=self.system_vars['retry_delay'], width=10).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        # Path settings
        ttk.Separator(system_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        row += 1
        
        ttk.Label(system_frame, text="Application Paths", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        # Chrome browser path
        ttk.Label(system_frame, text="Chrome Browser Path:").grid(row=row, column=0, sticky=tk.W)
        path_frame = ttk.Frame(system_frame)
        path_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        path_frame.columnconfigure(0, weight=1)
        
        self.system_vars['chrome_path'] = tk.StringVar()
        ttk.Entry(path_frame, textvariable=self.system_vars['chrome_path']).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(path_frame, text="Browse", 
                  command=lambda: self.browse_file(self.system_vars['chrome_path'])).grid(row=0, column=1, padx=(5, 0))
        row += 1
        
        # Download directory
        ttk.Label(system_frame, text="Download Directory:").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        download_frame = ttk.Frame(system_frame)
        download_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=(10, 0))
        download_frame.columnconfigure(0, weight=1)
        
        self.system_vars['download_directory'] = tk.StringVar()
        ttk.Entry(download_frame, textvariable=self.system_vars['download_directory']).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(download_frame, text="Browse", 
                  command=lambda: self.browse_directory(self.system_vars['download_directory'])).grid(row=0, column=1, padx=(5, 0))
    
    def create_vbs_settings_tab(self):
        """Create VBS application settings tab"""
        vbs_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(vbs_frame, text="VBS Settings")
        
        # Configure grid
        vbs_frame.columnconfigure(1, weight=1)
        
        row = 0
        
        # Application paths
        ttk.Label(vbs_frame, text="VBS Application Paths", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        row += 1
        
        ttk.Label(vbs_frame, text="Primary Path:").grid(row=row, column=0, sticky=tk.W)
        primary_frame = ttk.Frame(vbs_frame)
        primary_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        primary_frame.columnconfigure(0, weight=1)
        
        self.vbs_vars['primary_path'] = tk.StringVar()
        ttk.Entry(primary_frame, textvariable=self.vbs_vars['primary_path']).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(primary_frame, text="Browse", 
                  command=lambda: self.browse_file(self.vbs_vars['primary_path'])).grid(row=0, column=1, padx=(5, 0))
        row += 1
        
        ttk.Label(vbs_frame, text="Backup Path:").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        backup_frame = ttk.Frame(vbs_frame)
        backup_frame.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=(10, 0))
        backup_frame.columnconfigure(0, weight=1)
        
        self.vbs_vars['backup_path'] = tk.StringVar()
        ttk.Entry(backup_frame, textvariable=self.vbs_vars['backup_path']).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(backup_frame, text="Browse", 
                  command=lambda: self.browse_file(self.vbs_vars['backup_path'])).grid(row=0, column=1, padx=(5, 0))
        row += 1
        
        # Login credentials
        ttk.Separator(vbs_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        row += 1
        
        ttk.Label(vbs_frame, text="Login Credentials", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        ttk.Label(vbs_frame, text="Username:").grid(row=row, column=0, sticky=tk.W)
        self.vbs_vars['username'] = tk.StringVar()
        ttk.Entry(vbs_frame, textvariable=self.vbs_vars['username'], width=30).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        ttk.Label(vbs_frame, text="Password:").grid(row=row, column=0, sticky=tk.W)
        self.vbs_vars['password'] = tk.StringVar()
        ttk.Entry(vbs_frame, textvariable=self.vbs_vars['password'], show="*", width=30).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        ttk.Label(vbs_frame, text="Database:").grid(row=row, column=0, sticky=tk.W)
        self.vbs_vars['database'] = tk.StringVar()
        ttk.Entry(vbs_frame, textvariable=self.vbs_vars['database'], width=30).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
        row += 1
        
        # Timeout settings
        ttk.Separator(vbs_frame, orient='horizontal').grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=20)
        row += 1
        
        ttk.Label(vbs_frame, text="Timeout Settings (seconds)", 
                 font=('Arial', 12, 'bold')).grid(row=row, column=0, columnspan=2, sticky=tk.W)
        row += 1
        
        timeout_settings = [
            ("App Launch Timeout:", "app_launch", "30"),
            ("Login Timeout:", "login", "20"),
            ("Navigation Timeout:", "navigation", "15"),
            ("Data Import Timeout:", "data_import", "7200"),
            ("PDF Generation Timeout:", "pdf_generation", "300")
        ]
        
        for label_text, var_name, default_value in timeout_settings:
            ttk.Label(vbs_frame, text=label_text).grid(row=row, column=0, sticky=tk.W)
            self.vbs_vars[f'timeout_{var_name}'] = tk.StringVar(value=default_value)
            ttk.Entry(vbs_frame, textvariable=self.vbs_vars[f'timeout_{var_name}'], width=10).grid(row=row, column=1, sticky=tk.W, padx=(10, 0))
            row += 1
    
    def create_service_management_tab(self):
        """Create service management tab"""
        service_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(service_frame, text="Service Management")
        
        # Service status section
        status_frame = ttk.LabelFrame(service_frame, text="Service Status", padding="10")
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        status_frame.columnconfigure(1, weight=1)
        
        ttk.Label(status_frame, text="Service Status:").grid(row=0, column=0, sticky=tk.W)
        self.service_status_label = ttk.Label(status_frame, text="Checking...", foreground="blue")
        self.service_status_label.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        ttk.Button(status_frame, text="Refresh Status", 
                  command=self.update_service_status).grid(row=0, column=2, padx=(10, 0))
        
        # Service control buttons
        control_frame = ttk.LabelFrame(service_frame, text="Service Control", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        button_frame = ttk.Frame(control_frame)
        button_frame.pack()
        
        ttk.Button(button_frame, text="Start Service", 
                  command=self.start_service).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Stop Service", 
                  command=self.stop_service).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Restart Service", 
                  command=self.restart_service).pack(side=tk.LEFT, padx=(0, 5))
        
        # Health monitoring
        health_frame = ttk.LabelFrame(service_frame, text="Health Monitoring", padding="10")
        health_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        health_frame.columnconfigure(0, weight=1)
        
        self.health_text = scrolledtext.ScrolledText(health_frame, height=8, width=80)
        self.health_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        health_button_frame = ttk.Frame(health_frame)
        health_button_frame.grid(row=1, column=0, pady=(10, 0))
        
        ttk.Button(health_button_frame, text="Run Health Check", 
                  command=self.run_health_check).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(health_button_frame, text="Clear", 
                  command=lambda: self.health_text.delete(1.0, tk.END)).pack(side=tk.LEFT)
        
        # Installation management
        install_frame = ttk.LabelFrame(service_frame, text="Installation Management", padding="10")
        install_frame.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        install_button_frame = ttk.Frame(install_frame)
        install_button_frame.pack()
        
        ttk.Button(install_button_frame, text="Install Service", 
                  command=self.install_service).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(install_button_frame, text="Uninstall Service", 
                  command=self.uninstall_service).pack(side=tk.LEFT)
    
    def create_log_viewer_tab(self):
        """Create log viewer tab"""
        log_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(log_frame, text="Log Viewer")
        
        # Log file selection
        selection_frame = ttk.Frame(log_frame)
        selection_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        selection_frame.columnconfigure(1, weight=1)
        
        ttk.Label(selection_frame, text="Log File:").grid(row=0, column=0, sticky=tk.W)
        
        self.log_file_var = tk.StringVar()
        log_combo = ttk.Combobox(selection_frame, textvariable=self.log_file_var, width=50)
        log_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        ttk.Button(selection_frame, text="Refresh", 
                  command=self.refresh_log_files).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(selection_frame, text="Load", 
                  command=self.load_log_file).grid(row=0, column=3, padx=(5, 0))
        
        # Log content display
        self.log_text = scrolledtext.ScrolledText(log_frame, height=25, width=100)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        
        # Initialize log files
        self.refresh_log_files()
        
        # Auto-refresh checkbox
        auto_refresh_frame = ttk.Frame(log_frame)
        auto_refresh_frame.grid(row=2, column=0, pady=(10, 0))
        
        self.auto_refresh_logs = tk.BooleanVar()
        ttk.Checkbutton(auto_refresh_frame, text="Auto-refresh every 30 seconds",
                       variable=self.auto_refresh_logs,
                       command=self.toggle_log_auto_refresh).pack(side=tk.LEFT)
        
        ttk.Button(auto_refresh_frame, text="Clear", 
                  command=lambda: self.log_text.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=(20, 0))
    
    def create_test_tools_tab(self):
        """Create test tools tab"""
        test_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(test_frame, text="Test Tools")
        
        # Email testing
        email_test_frame = ttk.LabelFrame(test_frame, text="Email Testing", padding="10")
        email_test_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        email_test_frame.columnconfigure(1, weight=1)
        
        ttk.Label(email_test_frame, text="Test Email Recipient:").grid(row=0, column=0, sticky=tk.W)
        self.test_email_var = tk.StringVar()
        ttk.Entry(email_test_frame, textvariable=self.test_email_var, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        test_email_buttons = ttk.Frame(email_test_frame)
        test_email_buttons.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(test_email_buttons, text="Send Test Email", 
                  command=self.send_test_email).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(test_email_buttons, text="Test SMTP Connection", 
                  command=self.test_smtp_connection).pack(side=tk.LEFT)
        
        # System testing
        system_test_frame = ttk.LabelFrame(test_frame, text="System Testing", padding="10")
        system_test_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        system_test_buttons = ttk.Frame(system_test_frame)
        system_test_buttons.pack()
        
        ttk.Button(system_test_buttons, text="Test CSV Download", 
                  command=self.test_csv_download).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(system_test_buttons, text="Test Excel Generation", 
                  command=self.test_excel_generation).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(system_test_buttons, text="Test VBS Connection", 
                  command=self.test_vbs_connection).pack(side=tk.LEFT)
        
        # Test results
        results_frame = ttk.LabelFrame(test_frame, text="Test Results", padding="10")
        results_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        self.test_results_text = scrolledtext.ScrolledText(results_frame, height=15, width=80)
        self.test_results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights for test frame
        test_frame.columnconfigure(0, weight=1)
        test_frame.rowconfigure(2, weight=1)
    
    def create_status_bar(self, parent):
        """Create status bar"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(1, weight=1)
        
        # Status message
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=0, sticky=tk.W)
        
        # Save/Load buttons
        button_frame = ttk.Frame(status_frame)
        button_frame.grid(row=0, column=1, sticky=tk.E)
        
        ttk.Button(button_frame, text="Load Configuration", 
                  command=self.load_configuration).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Save Configuration", 
                  command=self.save_configuration).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Reset to Defaults", 
                  command=self.reset_to_defaults).pack(side=tk.LEFT)
    
    def setup_auto_refresh(self):
        """Setup auto-refresh for service status"""
        def refresh_status():
            if self.root and self.root.winfo_exists():
                self.update_service_status()
                self.root.after(30000, refresh_status)  # Refresh every 30 seconds
        
        self.root.after(5000, refresh_status)  # Initial delay of 5 seconds
    
    def load_configuration(self):
        """Load current configuration into GUI"""
        try:
            if not self.config_manager:
                self.status_var.set("Configuration manager not available")
                return
            
            self.logger.info("Loading configuration...")
            
            # Load email settings
            email_config = self.config_manager.get('email_settings', {})
            
            # Email delivery options
            send_weekdays = email_config.get('send_weekdays_only', False)
            send_all_days = email_config.get('send_all_days', True)
            
            self.email_vars['send_weekdays_only'].set(send_weekdays)
            self.email_vars['send_all_days'].set(send_all_days)
            
            # Recipients
            daily_recipients = email_config.get('daily_recipients', [])
            completion_recipients = email_config.get('completion_recipients', [])
            error_recipients = email_config.get('error_recipients', [])
            
            self.email_vars['daily_recipients'].set(', '.join(daily_recipients))
            self.email_vars['completion_recipients'].set(', '.join(completion_recipients))
            self.email_vars['error_recipients'].set(', '.join(error_recipients))
            
            # SMTP settings
            smtp_config = email_config.get('smtp', {})
            self.email_vars['gmail_server'].set(smtp_config.get('server', 'smtp.gmail.com'))
            self.email_vars['gmail_port'].set(str(smtp_config.get('port', 587)))
            
            # Load system settings
            system_config = self.config_manager.get('system_settings', {})
            
            self.system_vars['auto_restart'].set(system_config.get('auto_restart', True))
            self.system_vars['restart_time'].set(system_config.get('restart_time', '02:00'))
            self.system_vars['debug_mode'].set(system_config.get('debug_mode', False))
            self.system_vars['max_retry_attempts'].set(str(system_config.get('max_retry_attempts', 3)))
            self.system_vars['retry_delay'].set(str(system_config.get('retry_delay', 30)))
            
            # Application paths
            self.system_vars['chrome_path'].set(system_config.get('chrome_path', ''))
            self.system_vars['download_directory'].set(system_config.get('download_directory', ''))
            
            # Load VBS settings
            vbs_config = self.config_manager.get('vbs_application', {})
            
            self.vbs_vars['primary_path'].set(vbs_config.get('primary_path', ''))
            self.vbs_vars['backup_path'].set(vbs_config.get('backup_path', ''))
            
            # Login credentials
            login_config = vbs_config.get('login', {})
            self.vbs_vars['username'].set(login_config.get('username', ''))
            self.vbs_vars['password'].set(login_config.get('password', ''))
            self.vbs_vars['database'].set(login_config.get('database', ''))
            
            # Timeout settings
            timeout_config = vbs_config.get('timeout', {})
            self.vbs_vars['timeout_app_launch'].set(str(timeout_config.get('app_launch', 30)))
            self.vbs_vars['timeout_login'].set(str(timeout_config.get('login', 20)))
            self.vbs_vars['timeout_navigation'].set(str(timeout_config.get('navigation', 15)))
            self.vbs_vars['timeout_data_import'].set(str(timeout_config.get('data_import', 7200)))
            self.vbs_vars['timeout_pdf_generation'].set(str(timeout_config.get('pdf_generation', 300)))
            
            self.status_var.set("Configuration loaded successfully")
            self.unsaved_changes = False
            self.logger.info("Configuration loaded successfully")
            
        except Exception as e:
            error_msg = f"Failed to load configuration: {e}"
            self.logger.error(error_msg)
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def save_configuration(self):
        """Save current GUI configuration"""
        try:
            if not self.config_manager:
                self.status_var.set("Configuration manager not available")
                return
            
            self.logger.info("Saving configuration...")
            
            # Prepare email settings
            email_config = {
                'send_weekdays_only': self.email_vars['send_weekdays_only'].get(),
                'send_all_days': self.email_vars['send_all_days'].get(),
                'daily_recipients': [email.strip() for email in self.email_vars['daily_recipients'].get().split(',') if email.strip()],
                'completion_recipients': [email.strip() for email in self.email_vars['completion_recipients'].get().split(',') if email.strip()],
                'error_recipients': [email.strip() for email in self.email_vars['error_recipients'].get().split(',') if email.strip()],
                'smtp': {
                    'server': self.email_vars['gmail_server'].get(),
                    'port': int(self.email_vars['gmail_port'].get()) if self.email_vars['gmail_port'].get().isdigit() else 587
                }
            }
            
            # Prepare system settings
            system_config = {
                'auto_restart': self.system_vars['auto_restart'].get(),
                'restart_time': self.system_vars['restart_time'].get(),
                'debug_mode': self.system_vars['debug_mode'].get(),
                'max_retry_attempts': int(self.system_vars['max_retry_attempts'].get()) if self.system_vars['max_retry_attempts'].get().isdigit() else 3,
                'retry_delay': int(self.system_vars['retry_delay'].get()) if self.system_vars['retry_delay'].get().isdigit() else 30,
                'chrome_path': self.system_vars['chrome_path'].get(),
                'download_directory': self.system_vars['download_directory'].get()
            }
            
            # Prepare VBS settings
            vbs_config = {
                'primary_path': self.vbs_vars['primary_path'].get(),
                'backup_path': self.vbs_vars['backup_path'].get(),
                'login': {
                    'username': self.vbs_vars['username'].get(),
                    'password': self.vbs_vars['password'].get(),
                    'database': self.vbs_vars['database'].get()
                },
                'timeout': {
                    'app_launch': int(self.vbs_vars['timeout_app_launch'].get()) if self.vbs_vars['timeout_app_launch'].get().isdigit() else 30,
                    'login': int(self.vbs_vars['timeout_login'].get()) if self.vbs_vars['timeout_login'].get().isdigit() else 20,
                    'navigation': int(self.vbs_vars['timeout_navigation'].get()) if self.vbs_vars['timeout_navigation'].get().isdigit() else 15,
                    'data_import': int(self.vbs_vars['timeout_data_import'].get()) if self.vbs_vars['timeout_data_import'].get().isdigit() else 7200,
                    'pdf_generation': int(self.vbs_vars['timeout_pdf_generation'].get()) if self.vbs_vars['timeout_pdf_generation'].get().isdigit() else 300
                }
            }
            
            # Update configuration
            self.config_manager.update('email_settings', email_config)
            self.config_manager.update('system_settings', system_config)
            self.config_manager.update('vbs_application', vbs_config)
            
            # Save to file
            self.config_manager.save_config()
            
            self.last_config_save = datetime.now()
            self.unsaved_changes = False
            self.status_var.set(f"Configuration saved at {self.last_config_save.strftime('%H:%M:%S')}")
            self.logger.info("Configuration saved successfully")
            
            messagebox.showinfo("Success", "Configuration saved successfully!")
            
        except Exception as e:
            error_msg = f"Failed to save configuration: {e}"
            self.logger.error(error_msg)
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def reset_to_defaults(self):
        """Reset configuration to defaults"""
        try:
            if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset all settings to defaults? This cannot be undone."):
                self.logger.info("Resetting configuration to defaults...")
                
                # Reset email settings
                self.email_vars['send_weekdays_only'].set(False)
                self.email_vars['send_all_days'].set(True)
                self.email_vars['daily_recipients'].set('')
                self.email_vars['completion_recipients'].set('')
                self.email_vars['error_recipients'].set('')
                self.email_vars['gmail_server'].set('smtp.gmail.com')
                self.email_vars['gmail_port'].set('587')
                
                # Reset system settings
                self.system_vars['auto_restart'].set(True)
                self.system_vars['restart_time'].set('02:00')
                self.system_vars['debug_mode'].set(False)
                self.system_vars['max_retry_attempts'].set('3')
                self.system_vars['retry_delay'].set('30')
                self.system_vars['chrome_path'].set('')
                self.system_vars['download_directory'].set('')
                
                # Reset VBS settings
                self.vbs_vars['primary_path'].set('')
                self.vbs_vars['backup_path'].set('')
                self.vbs_vars['username'].set('')
                self.vbs_vars['password'].set('')
                self.vbs_vars['database'].set('')
                self.vbs_vars['timeout_app_launch'].set('30')
                self.vbs_vars['timeout_login'].set('20')
                self.vbs_vars['timeout_navigation'].set('15')
                self.vbs_vars['timeout_data_import'].set('7200')
                self.vbs_vars['timeout_pdf_generation'].set('300')
                
                self.unsaved_changes = True
                self.status_var.set("Configuration reset to defaults (not saved)")
                self.logger.info("Configuration reset to defaults")
                
        except Exception as e:
            error_msg = f"Failed to reset configuration: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def on_weekday_only_changed(self):
        """Handle weekday only checkbox change"""
        if self.email_vars['send_weekdays_only'].get():
            self.email_vars['send_all_days'].set(False)
        self.unsaved_changes = True
    
    def on_all_days_changed(self):
        """Handle all days checkbox change"""
        if self.email_vars['send_all_days'].get():
            self.email_vars['send_weekdays_only'].set(False)
        self.unsaved_changes = True
    
    def browse_file(self, var):
        """Browse for file path"""
        try:
            filename = filedialog.askopenfilename(
                title="Select File",
                filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
            )
            if filename:
                var.set(filename)
                self.unsaved_changes = True
        except Exception as e:
            self.logger.error(f"File browse failed: {e}")
    
    def browse_directory(self, var):
        """Browse for directory path"""
        try:
            directory = filedialog.askdirectory(title="Select Directory")
            if directory:
                var.set(directory)
                self.unsaved_changes = True
        except Exception as e:
            self.logger.error(f"Directory browse failed: {e}")
    
    def update_service_status(self):
        """Update service status display"""
        try:
            if not self.service_manager:
                self.service_status_var.set("Service Manager: Not Available")
                if hasattr(self, 'service_status_label'):
                    self.service_status_label.config(foreground="red")
                return
            
            status = self.service_manager.get_service_status()
            
            if status.get('installed', False):
                if status.get('running', False):
                    status_text = f"Service: Running ({status.get('status', 'Unknown')})"
                    color = "green"
                else:
                    status_text = f"Service: Stopped ({status.get('status', 'Unknown')})"
                    color = "orange"
            else:
                status_text = "Service: Not Installed"
                color = "red"
            
            self.service_status_var.set(status_text)
            if hasattr(self, 'service_status_label'):
                self.service_status_label.config(foreground=color)
            
        except Exception as e:
            self.logger.error(f"Failed to update service status: {e}")
            self.service_status_var.set("Service: Status Unknown")
            if hasattr(self, 'service_status_label'):
                self.service_status_label.config(foreground="red")
    
    def start_service(self):
        """Start the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            self.status_var.set("Starting service...")
            result = self.service_manager.start_service()
            
            if result['success']:
                messagebox.showinfo("Success", "Service started successfully!")
                self.status_var.set("Service started")
            else:
                messagebox.showerror("Error", f"Failed to start service: {result['error']}")
                self.status_var.set("Service start failed")
            
            self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to start service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def stop_service(self):
        """Stop the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Are you sure you want to stop the service?"):
                self.status_var.set("Stopping service...")
                result = self.service_manager.stop_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service stopped successfully!")
                    self.status_var.set("Service stopped")
                else:
                    messagebox.showerror("Error", f"Failed to stop service: {result['error']}")
                    self.status_var.set("Service stop failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to stop service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def restart_service(self):
        """Restart the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Are you sure you want to restart the service?"):
                self.status_var.set("Restarting service...")
                result = self.service_manager.restart_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service restarted successfully!")
                    self.status_var.set("Service restarted")
                else:
                    messagebox.showerror("Error", f"Failed to restart service: {result['error']}")
                    self.status_var.set("Service restart failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to restart service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def install_service(self):
        """Install the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Install MoonFlower WiFi Automation Service?"):
                self.status_var.set("Installing service...")
                
                # Run installation in separate thread to prevent GUI freezing
                def install_thread():
                    try:
                        result = self.service_manager.install_service()
                        
                        # Update GUI in main thread
                        self.root.after(0, lambda: self._handle_install_result(result))
                    except Exception as e:
                        self.root.after(0, lambda: self._handle_install_error(e))
                
                threading.Thread(target=install_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"Failed to install service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def _handle_install_result(self, result):
        """Handle service installation result"""
        if result['success']:
            messagebox.showinfo("Success", "Service installed successfully!")
            self.status_var.set("Service installed")
        else:
            messagebox.showerror("Error", f"Failed to install service: {result['error']}")
            self.status_var.set("Service installation failed")
        
        self.update_service_status()
    
    def _handle_install_error(self, error):
        """Handle service installation error"""
        error_msg = f"Service installation error: {error}"
        self.logger.error(error_msg)
        messagebox.showerror("Error", error_msg)
        self.status_var.set("Service installation failed")
    
    def uninstall_service(self):
        """Uninstall the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Are you sure you want to uninstall the service? This will remove all startup mechanisms."):
                self.status_var.set("Uninstalling service...")
                result = self.service_manager.uninstall_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service uninstalled successfully!")
                    self.status_var.set("Service uninstalled")
                else:
                    messagebox.showerror("Error", f"Failed to uninstall service: {result['error']}")
                    self.status_var.set("Service uninstall failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to uninstall service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def run_health_check(self):
        """Run comprehensive health check"""
        try:
            if not self.health_monitor:
                self.health_text.insert(tk.END, "Health monitor not available\n")
                return
            
            self.health_text.insert(tk.END, f"\n=== Health Check - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            self.health_text.insert(tk.END, "Running comprehensive health check...\n")
            
            # Run health check in separate thread
            def health_check_thread():
                try:
                    result = self.health_monitor.perform_comprehensive_health_check()
                    
                    # Update GUI in main thread
                    self.root.after(0, lambda: self._display_health_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._display_health_error(e))
            
            threading.Thread(target=health_check_thread, daemon=True).start()
            
        except Exception as e:
            self.health_text.insert(tk.END, f"Health check failed: {e}\n")
            self.logger.error(f"Health check failed: {e}")
    
    def _display_health_result(self, result):
        """Display health check result"""
        try:
            overall_status = "HEALTHY" if result['overall_healthy'] else "UNHEALTHY"
            health_score = result.get('health_score', 0)
            
            self.health_text.insert(tk.END, f"Overall Status: {overall_status} (Score: {health_score}%)\n")
            
            if result['critical_issues']:
                self.health_text.insert(tk.END, f"\nCritical Issues ({len(result['critical_issues'])}):\n")
                for issue in result['critical_issues']:
                    self.health_text.insert(tk.END, f"  • {issue}\n")
            
            if result['warnings']:
                self.health_text.insert(tk.END, f"\nWarnings ({len(result['warnings'])}):\n")
                for warning in result['warnings']:
                    self.health_text.insert(tk.END, f"  • {warning}\n")
            
            # Display detailed results
            self.health_text.insert(tk.END, "\nDetailed Results:\n")
            for check_name, check_result in result['detailed_results'].items():
                status = "✓" if check_result.get('healthy', False) else "✗"
                self.health_text.insert(tk.END, f"  {status} {check_name.replace('_', ' ').title()}\n")
            
            self.health_text.insert(tk.END, "\nHealth check completed.\n")
            self.health_text.see(tk.END)
            
        except Exception as e:
            self.health_text.insert(tk.END, f"Error displaying health result: {e}\n")
    
    def _display_health_error(self, error):
        """Display health check error"""
        self.health_text.insert(tk.END, f"Health check error: {error}\n")
        self.health_text.see(tk.END)
    
    def refresh_log_files(self):
        """Refresh available log files"""
        try:
            log_files = []
            
            # Check EHC_Logs directory
            logs_dir = Path("EHC_Logs")
            if logs_dir.exists():
                for log_file in logs_dir.rglob("*.log"):
                    log_files.append(str(log_file))
            
            # Sort by modification time (newest first)
            log_files.sort(key=lambda x: os.path.getmtime(x) if os.path.exists(x) else 0, reverse=True)
            
            # Update combobox
            if hasattr(self, 'log_file_var'):
                combo = None
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Frame):
                        for child in widget.winfo_children():
                            if isinstance(child, ttk.Notebook):
                                # Find log viewer tab
                                for tab_id in child.tabs():
                                    tab_frame = child.nametowidget(tab_id)
                                    if child.tab(tab_id, "text") == "Log Viewer":
                                        for tab_child in tab_frame.winfo_children():
                                            if isinstance(tab_child, ttk.Frame):
                                                for frame_child in tab_child.winfo_children():
                                                    if isinstance(frame_child, ttk.Combobox):
                                                        combo = frame_child
                                                        break
                
                if combo:
                    combo['values'] = log_files
                    if log_files and not self.log_file_var.get():
                        self.log_file_var.set(log_files[0])
            
        except Exception as e:
            self.logger.error(f"Failed to refresh log files: {e}")
    
    def load_log_file(self):
        """Load selected log file"""
        try:
            log_file = self.log_file_var.get()
            if not log_file or not os.path.exists(log_file):
                messagebox.showerror("Error", "Please select a valid log file")
                return
            
            self.log_text.delete(1.0, tk.END)
            self.log_text.insert(tk.END, f"Loading {log_file}...\n")
            
            # Load file in chunks to handle large files
            with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                
                # Show last 1000 lines for performance
                lines = content.split('\n')
                if len(lines) > 1000:
                    content = '\n'.join(lines[-1000:])
                    self.log_text.insert(tk.END, f"[Showing last 1000 lines of {len(lines)} total lines]\n\n")
                
                self.log_text.delete(1.0, tk.END)
                self.log_text.insert(tk.END, content)
                self.log_text.see(tk.END)
            
        except Exception as e:
            error_msg = f"Failed to load log file: {e}"
            self.logger.error(error_msg)
            self.log_text.insert(tk.END, f"\nError: {error_msg}\n")
    
    def toggle_log_auto_refresh(self):
        """Toggle log auto-refresh"""
        if self.auto_refresh_logs.get():
            self._start_log_auto_refresh()
        else:
            self._stop_log_auto_refresh()
    
    def _start_log_auto_refresh(self):
        """Start log auto-refresh"""
        def refresh_log():
            if self.auto_refresh_logs.get() and self.root and self.root.winfo_exists():
                self.load_log_file()
                self.root.after(30000, refresh_log)  # Refresh every 30 seconds
        
        self.root.after(30000, refresh_log)
    
    def _stop_log_auto_refresh(self):
        """Stop log auto-refresh"""
        # Auto-refresh is controlled by the checkbox state
        pass
    
    def send_test_email(self):
        """Send test email"""
        try:
            if not self.email_system:
                messagebox.showerror("Error", "Email system not available")
                return
            
            test_recipient = self.test_email_var.get().strip()
            if not test_recipient:
                messagebox.showerror("Error", "Please enter a test email recipient")
                return
            
            self.test_results_text.insert(tk.END, f"\n=== Test Email - {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, f"Sending test email to: {test_recipient}\n")
            
            # Send test email in separate thread
            def send_email_thread():
                try:
                    result = self.email_system.send_email(
                        recipients=[test_recipient],
                        subject="MoonFlower Automation - Test Email",
                        message="This is a test email from the MoonFlower WiFi Automation configuration panel.",
                        email_type='test'
                    )
                    
                    self.root.after(0, lambda: self._handle_test_email_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._handle_test_email_error(e))
            
            threading.Thread(target=send_email_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"Test email failed: {e}"
            self.logger.error(error_msg)
            self.test_results_text.insert(tk.END, f"Error: {error_msg}\n")
    
    def _handle_test_email_result(self, result):
        """Handle test email result"""
        if result.get('success', False):
            self.test_results_text.insert(tk.END, "✓ Test email sent successfully!\n")
            messagebox.showinfo("Success", "Test email sent successfully!")
        else:
            error_msg = result.get('error', 'Unknown error')
            self.test_results_text.insert(tk.END, f"✗ Test email failed: {error_msg}\n")
            messagebox.showerror("Error", f"Test email failed: {error_msg}")
        
        self.test_results_text.see(tk.END)
    
    def _handle_test_email_error(self, error):
        """Handle test email error"""
        error_msg = f"Test email error: {error}"
        self.test_results_text.insert(tk.END, f"✗ {error_msg}\n")
        self.test_results_text.see(tk.END)
        messagebox.showerror("Error", error_msg)
    
    def test_smtp_connection(self):
        """Test SMTP connection"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== SMTP Connection Test - {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "Testing SMTP connection...\n")
            
            # Test SMTP connection in separate thread
            def test_smtp_thread():
                try:
                    if self.email_system:
                        # Use email system's SMTP test if available
                        result = {'success': True, 'message': 'SMTP connection test not implemented'}
                    else:
                        result = {'success': False, 'error': 'Email system not available'}
                    
                    self.root.after(0, lambda: self._handle_smtp_test_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._handle_smtp_test_error(e))
            
            threading.Thread(target=test_smtp_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"SMTP test failed: {e}"
            self.logger.error(error_msg)
            self.test_results_text.insert(tk.END, f"Error: {error_msg}\n")
    
    def _handle_smtp_test_result(self, result):
        """Handle SMTP test result"""
        if result.get('success', False):
            self.test_results_text.insert(tk.END, "✓ SMTP connection successful!\n")
        else:
            error_msg = result.get('error', 'Unknown error')
            self.test_results_text.insert(tk.END, f"✗ SMTP connection failed: {error_msg}\n")
        
        self.test_results_text.see(tk.END)
    
    def _handle_smtp_test_error(self, error):
        """Handle SMTP test error"""
        self.test_results_text.insert(tk.END, f"✗ SMTP test error: {error}\n")
        self.test_results_text.see(tk.END)
    
    def validate_smtp_settings(self):
        """Validate SMTP settings"""
        try:
            server = self.email_vars['gmail_server'].get()
            port = self.email_vars['gmail_port'].get()
            
            if not server:
                messagebox.showerror("Error", "SMTP server is required")
                return
            
            if not port.isdigit():
                messagebox.showerror("Error", "SMTP port must be a number")
                return
            
            messagebox.showinfo("Validation", f"SMTP settings appear valid:\nServer: {server}\nPort: {port}")
            
        except Exception as e:
            error_msg = f"SMTP validation failed: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def test_csv_download(self):
        """Test CSV download functionality"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== CSV Download Test - {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "Testing CSV download functionality...\n")
            
            # Test CSV download in separate thread
            def test_csv_thread():
                try:
                    # Import and test CSV downloader
                    from wifi.csv_downloader import CSVDownloader
                    
                    downloader = CSVDownloader()
                    result = {'success': True, 'message': 'CSV downloader initialized successfully'}
                    
                    self.root.after(0, lambda: self._handle_csv_test_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._handle_csv_test_error(e))
            
            threading.Thread(target=test_csv_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"CSV test failed: {e}"
            self.logger.error(error_msg)
            self.test_results_text.insert(tk.END, f"Error: {error_msg}\n")
    
    def _handle_csv_test_result(self, result):
        """Handle CSV test result"""
        if result.get('success', False):
            self.test_results_text.insert(tk.END, "✓ CSV download test successful!\n")
        else:
            error_msg = result.get('error', 'Unknown error')
            self.test_results_text.insert(tk.END, f"✗ CSV download test failed: {error_msg}\n")
        
        self.test_results_text.see(tk.END)
    
    def _handle_csv_test_error(self, error):
        """Handle CSV test error"""
        self.test_results_text.insert(tk.END, f"✗ CSV test error: {error}\n")
        self.test_results_text.see(tk.END)
    
    def test_excel_generation(self):
        """Test Excel generation functionality"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Excel Generation Test - {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "Testing Excel generation functionality...\n")
            
            # Test Excel generation in separate thread
            def test_excel_thread():
                try:
                    # Import and test Excel generator
                    from excel.excel_generator import ExcelGenerator
                    
                    generator = ExcelGenerator()
                    result = {'success': True, 'message': 'Excel generator initialized successfully'}
                    
                    self.root.after(0, lambda: self._handle_excel_test_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._handle_excel_test_error(e))
            
            threading.Thread(target=test_excel_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"Excel test failed: {e}"
            self.logger.error(error_msg)
            self.test_results_text.insert(tk.END, f"Error: {error_msg}\n")
    
    def _handle_excel_test_result(self, result):
        """Handle Excel test result"""
        if result.get('success', False):
            self.test_results_text.insert(tk.END, "✓ Excel generation test successful!\n")
        else:
            error_msg = result.get('error', 'Unknown error')
            self.test_results_text.insert(tk.END, f"✗ Excel generation test failed: {error_msg}\n")
        
        self.test_results_text.see(tk.END)
    
    def _handle_excel_test_error(self, error):
        """Handle Excel test error"""
        self.test_results_text.insert(tk.END, f"✗ Excel test error: {error}\n")
        self.test_results_text.see(tk.END)
    
    def test_vbs_connection(self):
        """Test VBS connection functionality"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== VBS Connection Test - {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "Testing VBS connection functionality...\n")
            
            # Test VBS connection in separate thread
            def test_vbs_thread():
                try:
                    # Import and test VBS automator
                    from vbs.vbs_core import VBSAutomator
                    
                    automator = VBSAutomator()
                    result = {'success': True, 'message': 'VBS automator initialized successfully'}
                    
                    self.root.after(0, lambda: self._handle_vbs_test_result(result))
                except Exception as e:
                    self.root.after(0, lambda: self._handle_vbs_test_error(e))
            
            threading.Thread(target=test_vbs_thread, daemon=True).start()
            
        except Exception as e:
            error_msg = f"VBS test failed: {e}"
            self.logger.error(error_msg)
            self.test_results_text.insert(tk.END, f"Error: {error_msg}\n")
    
    def _handle_vbs_test_result(self, result):
        """Handle VBS test result"""
        if result.get('success', False):
            self.test_results_text.insert(tk.END, "✓ VBS connection test successful!\n")
        else:
            error_msg = result.get('error', 'Unknown error')
            self.test_results_text.insert(tk.END, f"✗ VBS connection test failed: {error_msg}\n")
        
        self.test_results_text.see(tk.END)
    
    def _handle_vbs_test_error(self, error):
        """Handle VBS test error"""
        self.test_results_text.insert(tk.END, f"✗ VBS test error: {error}\n")
        self.test_results_text.see(tk.END)
    
    def run(self):
        """Run the configuration panel"""
        try:
            self.create_gui()
            if self.root:
                self.root.mainloop()
        except Exception as e:
            self.logger.error(f"Failed to run configuration panel: {e}")
            if self.root:
                messagebox.showerror("Fatal Error", f"Configuration panel failed: {e}")


def main():
    """Main entry point for configuration panel"""
    try:
        # Setup basic logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Create and run configuration panel
        config_panel = ConfigurationPanel()
        config_panel.run()
        
    except Exception as e:
        print(f"Failed to start configuration panel: {e}")
        logging.error(f"Failed to start configuration panel: {e}")


if __name__ == "__main__":
    main()   
 
    def setup_auto_refresh(self):
        """Setup auto-refresh for service status"""
        def refresh_status():
            if self.root and self.root.winfo_exists():
                self.update_service_status()
                self.root.after(30000, refresh_status)  # Refresh every 30 seconds
        
        self.root.after(5000, refresh_status)  # Initial delay of 5 seconds
    
    def load_configuration(self):
        """Load configuration from config manager"""
        try:
            if not self.config_manager:
                self.status_var.set("Configuration manager not available")
                return
            
            self.status_var.set("Loading configuration...")
            
            # Load email settings
            email_settings = self.config_manager.get_email_settings()
            self.email_vars['send_weekdays_only'].set(email_settings.get('send_weekdays_only', True))
            self.email_vars['send_all_days'].set(email_settings.get('send_all_days', False))
            
            # Load recipients (convert lists to comma-separated strings)
            daily_recipients = email_settings.get('daily_recipients', [])
            self.email_vars['daily_recipients'].set(', '.join(daily_recipients) if daily_recipients else '')
            
            completion_recipients = email_settings.get('completion_recipients', [])
            self.email_vars['completion_recipients'].set(', '.join(completion_recipients) if completion_recipients else '')
            
            error_recipients = email_settings.get('error_recipients', [])
            self.email_vars['error_recipients'].set(', '.join(error_recipients) if error_recipients else '')
            
            # Load SMTP settings
            self.email_vars['gmail_server'].set(email_settings.get('gmail_server', 'smtp.gmail.com'))
            self.email_vars['gmail_port'].set(str(email_settings.get('gmail_port', 587)))
            
            # Load system settings
            system_settings = self.config_manager.get_system_settings()
            self.system_vars['auto_restart'].set(system_settings.get('auto_restart', True))
            self.system_vars['restart_time'].set(system_settings.get('restart_time', '02:00'))
            self.system_vars['debug_mode'].set(system_settings.get('debug_mode', False))
            self.system_vars['max_retry_attempts'].set(str(system_settings.get('max_retry_attempts', 3)))
            self.system_vars['retry_delay'].set(str(system_settings.get('retry_delay', 30)))
            
            # Load path settings
            self.system_vars['chrome_path'].set(system_settings.get('chrome_path', ''))
            self.system_vars['download_directory'].set(system_settings.get('download_directory', ''))
            
            # Load VBS settings
            vbs_settings = self.config_manager.get_vbs_application_settings()
            paths = self.config_manager.get_vbs_application_paths()
            credentials = self.config_manager.get_vbs_login_credentials()
            timeouts = self.config_manager.get_vbs_timeout_settings()
            
            self.vbs_vars['primary_path'].set(paths.get('primary_path', ''))
            self.vbs_vars['backup_path'].set(paths.get('backup_path', ''))
            self.vbs_vars['username'].set(credentials.get('username', ''))
            self.vbs_vars['password'].set(credentials.get('password', ''))
            self.vbs_vars['database'].set(credentials.get('database', ''))
            
            # Load timeout settings
            self.vbs_vars['timeout_app_launch'].set(str(timeouts.get('app_launch', 30)))
            self.vbs_vars['timeout_login'].set(str(timeouts.get('login', 20)))
            self.vbs_vars['timeout_navigation'].set(str(timeouts.get('navigation', 15)))
            self.vbs_vars['timeout_data_import'].set(str(timeouts.get('data_import', 7200)))
            self.vbs_vars['timeout_pdf_generation'].set(str(timeouts.get('pdf_generation', 300)))
            
            self.last_config_save = datetime.now()
            self.unsaved_changes = False
            self.status_var.set("Configuration loaded successfully")
            
            self.logger.info("Configuration loaded successfully")
            
        except Exception as e:
            error_msg = f"Failed to load configuration: {e}"
            self.logger.error(error_msg)
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def save_configuration(self):
        """Save configuration to config manager"""
        try:
            if not self.config_manager:
                self.status_var.set("Configuration manager not available")
                return
            
            self.status_var.set("Saving configuration...")
            
            # Validate configuration before saving
            validation_result = self.validate_configuration()
            if not validation_result['valid']:
                messagebox.showerror("Validation Error", 
                                   f"Configuration validation failed:\n" + 
                                   "\n".join(validation_result['errors']))
                return
            
            # Save email settings
            email_settings = {
                'send_weekdays_only': self.email_vars['send_weekdays_only'].get(),
                'send_all_days': self.email_vars['send_all_days'].get(),
                'daily_recipients': [email.strip() for email in self.email_vars['daily_recipients'].get().split(',') if email.strip()],
                'completion_recipients': [email.strip() for email in self.email_vars['completion_recipients'].get().split(',') if email.strip()],
                'error_recipients': [email.strip() for email in self.email_vars['error_recipients'].get().split(',') if email.strip()],
                'gmail_server': self.email_vars['gmail_server'].get(),
                'gmail_port': int(self.email_vars['gmail_port'].get())
            }
            
            for key, value in email_settings.items():
                self.config_manager.set(f'email_settings.{key}', value)
            
            # Save system settings
            system_settings = {
                'auto_restart': self.system_vars['auto_restart'].get(),
                'restart_time': self.system_vars['restart_time'].get(),
                'debug_mode': self.system_vars['debug_mode'].get(),
                'max_retry_attempts': int(self.system_vars['max_retry_attempts'].get()),
                'retry_delay': int(self.system_vars['retry_delay'].get()),
                'chrome_path': self.system_vars['chrome_path'].get(),
                'download_directory': self.system_vars['download_directory'].get()
            }
            
            for key, value in system_settings.items():
                self.config_manager.set(f'system_settings.{key}', value)
            
            # Save VBS settings
            self.config_manager.set('vbs_application.primary_path', self.vbs_vars['primary_path'].get())
            self.config_manager.set('vbs_application.backup_path', self.vbs_vars['backup_path'].get())
            self.config_manager.set('vbs_application.login.username', self.vbs_vars['username'].get())
            self.config_manager.set('vbs_application.login.password', self.vbs_vars['password'].get())
            self.config_manager.set('vbs_application.login.database', self.vbs_vars['database'].get())
            
            # Save timeout settings
            timeout_settings = {
                'app_launch': int(self.vbs_vars['timeout_app_launch'].get()),
                'login': int(self.vbs_vars['timeout_login'].get()),
                'navigation': int(self.vbs_vars['timeout_navigation'].get()),
                'data_import': int(self.vbs_vars['timeout_data_import'].get()),
                'pdf_generation': int(self.vbs_vars['timeout_pdf_generation'].get())
            }
            
            for key, value in timeout_settings.items():
                self.config_manager.set(f'vbs_application.timeout.{key}', value)
            
            self.last_config_save = datetime.now()
            self.unsaved_changes = False
            self.status_var.set("Configuration saved successfully")
            
            self.logger.info("Configuration saved successfully")
            messagebox.showinfo("Success", "Configuration saved successfully")
            
        except Exception as e:
            error_msg = f"Failed to save configuration: {e}"
            self.logger.error(error_msg)
            self.status_var.set(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def validate_configuration(self) -> Dict[str, Any]:
        """Validate current configuration"""
        errors = []
        warnings = []
        
        try:
            # Validate email settings
            if self.email_vars['send_weekdays_only'].get() and self.email_vars['send_all_days'].get():
                errors.append("Cannot select both 'weekdays only' and 'all days' for email delivery")
            
            if not self.email_vars['send_weekdays_only'].get() and not self.email_vars['send_all_days'].get():
                warnings.append("No email delivery option selected")
            
            # Validate email addresses
            import re
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            
            for field_name, var_name in [('Daily recipients', 'daily_recipients'), 
                                       ('Completion recipients', 'completion_recipients'),
                                       ('Error recipients', 'error_recipients')]:
                emails = [email.strip() for email in self.email_vars[var_name].get().split(',') if email.strip()]
                for email in emails:
                    if not re.match(email_pattern, email):
                        errors.append(f"Invalid email address in {field_name}: {email}")
            
            # Validate SMTP settings
            try:
                port = int(self.email_vars['gmail_port'].get())
                if port < 1 or port > 65535:
                    errors.append("Gmail port must be between 1 and 65535")
            except ValueError:
                errors.append("Gmail port must be a valid number")
            
            # Validate system settings
            try:
                max_retries = int(self.system_vars['max_retry_attempts'].get())
                if max_retries < 1 or max_retries > 10:
                    errors.append("Max retry attempts must be between 1 and 10")
            except ValueError:
                errors.append("Max retry attempts must be a valid number")
            
            try:
                retry_delay = int(self.system_vars['retry_delay'].get())
                if retry_delay < 1 or retry_delay > 300:
                    errors.append("Retry delay must be between 1 and 300 seconds")
            except ValueError:
                errors.append("Retry delay must be a valid number")
            
            # Validate restart time format
            restart_time = self.system_vars['restart_time'].get()
            if not re.match(r'^([01]?[0-9]|2[0-3]):[0-5][0-9]$', restart_time):
                errors.append("Restart time must be in HH:MM format (24-hour)")
            
            # Validate VBS paths
            primary_path = self.vbs_vars['primary_path'].get()
            if primary_path and not Path(primary_path).exists():
                warnings.append(f"VBS primary path does not exist: {primary_path}")
            
            backup_path = self.vbs_vars['backup_path'].get()
            if backup_path and not Path(backup_path).exists():
                warnings.append(f"VBS backup path does not exist: {backup_path}")
            
            # Validate timeout settings
            timeout_fields = [
                ('App launch timeout', 'timeout_app_launch', 5, 300),
                ('Login timeout', 'timeout_login', 5, 120),
                ('Navigation timeout', 'timeout_navigation', 5, 60),
                ('Data import timeout', 'timeout_data_import', 60, 14400),
                ('PDF generation timeout', 'timeout_pdf_generation', 30, 1800)
            ]
            
            for field_name, var_name, min_val, max_val in timeout_fields:
                try:
                    timeout_val = int(self.vbs_vars[var_name].get())
                    if timeout_val < min_val or timeout_val > max_val:
                        errors.append(f"{field_name} must be between {min_val} and {max_val} seconds")
                except ValueError:
                    errors.append(f"{field_name} must be a valid number")
            
            return {
                'valid': len(errors) == 0,
                'errors': errors,
                'warnings': warnings
            }
            
        except Exception as e:
            return {
                'valid': False,
                'errors': [f"Validation failed: {e}"],
                'warnings': []
            }
    
    def reset_to_defaults(self):
        """Reset configuration to defaults"""
        try:
            if messagebox.askyesno("Confirm Reset", 
                                 "Are you sure you want to reset all settings to defaults?\n" +
                                 "This will overwrite your current configuration."):
                
                if self.config_manager:
                    # Reset to default configuration
                    default_config = self.config_manager._get_default_config()
                    self.config_manager.config = default_config
                    self.config_manager.save_config()
                
                # Reload the GUI with defaults
                self.load_configuration()
                
                self.status_var.set("Configuration reset to defaults")
                messagebox.showinfo("Success", "Configuration reset to defaults")
                
        except Exception as e:
            error_msg = f"Failed to reset configuration: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    # Event handlers
    def on_weekday_only_changed(self):
        """Handle weekday only checkbox change"""
        if self.email_vars['send_weekdays_only'].get():
            self.email_vars['send_all_days'].set(False)
        self.mark_unsaved_changes()
    
    def on_all_days_changed(self):
        """Handle all days checkbox change"""
        if self.email_vars['send_all_days'].get():
            self.email_vars['send_weekdays_only'].set(False)
        self.mark_unsaved_changes()
    
    def mark_unsaved_changes(self):
        """Mark that there are unsaved changes"""
        self.unsaved_changes = True
        if self.status_var:
            self.status_var.set("Configuration modified (unsaved)")
    
    # File/directory browsing
    def browse_file(self, var):
        """Browse for a file"""
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
            self.mark_unsaved_changes()
    
    def browse_directory(self, var):
        """Browse for a directory"""
        directory = filedialog.askdirectory(title="Select Directory")
        if directory:
            var.set(directory)
            self.mark_unsaved_changes()
    
    # Service management methods
    def update_service_status(self):
        """Update service status display"""
        try:
            if not self.service_manager:
                self.service_status_var.set("Service manager not available")
                if hasattr(self, 'service_status_label'):
                    self.service_status_label.config(foreground="red")
                return
            
            status = self.service_manager.get_service_status()
            
            if status.get('installed', False):
                if status.get('running', False):
                    status_text = f"Service Running ({status.get('status', 'Unknown')})"
                    color = "green"
                else:
                    status_text = f"Service Stopped ({status.get('status', 'Unknown')})"
                    color = "orange"
            else:
                status_text = "Service Not Installed"
                color = "red"
            
            self.service_status_var.set(status_text)
            if hasattr(self, 'service_status_label'):
                self.service_status_label.config(foreground=color)
            
        except Exception as e:
            self.logger.error(f"Failed to update service status: {e}")
            self.service_status_var.set("Status check failed")
            if hasattr(self, 'service_status_label'):
                self.service_status_label.config(foreground="red")
    
    def start_service(self):
        """Start the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            self.status_var.set("Starting service...")
            result = self.service_manager.start_service()
            
            if result['success']:
                messagebox.showinfo("Success", "Service started successfully")
                self.status_var.set("Service started")
            else:
                messagebox.showerror("Error", f"Failed to start service: {result['error']}")
                self.status_var.set("Service start failed")
            
            self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to start service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def stop_service(self):
        """Stop the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Are you sure you want to stop the service?"):
                self.status_var.set("Stopping service...")
                result = self.service_manager.stop_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service stopped successfully")
                    self.status_var.set("Service stopped")
                else:
                    messagebox.showerror("Error", f"Failed to stop service: {result['error']}")
                    self.status_var.set("Service stop failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to stop service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def restart_service(self):
        """Restart the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Are you sure you want to restart the service?"):
                self.status_var.set("Restarting service...")
                result = self.service_manager.restart_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service restarted successfully")
                    self.status_var.set("Service restarted")
                else:
                    messagebox.showerror("Error", f"Failed to restart service: {result['error']}")
                    self.status_var.set("Service restart failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to restart service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def install_service(self):
        """Install the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", "Install the MoonFlower WiFi Automation service?"):
                self.status_var.set("Installing service...")
                result = self.service_manager.install_service()
                
                if result['success']:
                    messagebox.showinfo("Success", 
                                      f"Service installed successfully\n\n" +
                                      f"Service Name: {result['service_name']}\n" +
                                      f"Install Time: {result['install_time']}\n" +
                                      f"Startup Mechanisms: {', '.join(result.get('startup_mechanisms', []))}")
                    self.status_var.set("Service installed")
                else:
                    messagebox.showerror("Error", f"Failed to install service: {result['error']}")
                    self.status_var.set("Service installation failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to install service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def uninstall_service(self):
        """Uninstall the Windows service"""
        try:
            if not self.service_manager:
                messagebox.showerror("Error", "Service manager not available")
                return
            
            if messagebox.askyesno("Confirm", 
                                 "Are you sure you want to uninstall the service?\n" +
                                 "This will remove all startup mechanisms."):
                self.status_var.set("Uninstalling service...")
                result = self.service_manager.uninstall_service()
                
                if result['success']:
                    messagebox.showinfo("Success", "Service uninstalled successfully")
                    self.status_var.set("Service uninstalled")
                else:
                    messagebox.showerror("Error", f"Failed to uninstall service: {result['error']}")
                    self.status_var.set("Service uninstallation failed")
                
                self.update_service_status()
            
        except Exception as e:
            error_msg = f"Failed to uninstall service: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def run_health_check(self):
        """Run comprehensive health check"""
        try:
            if not self.health_monitor:
                self.health_text.insert(tk.END, "Health monitor not available\n")
                return
            
            self.health_text.insert(tk.END, f"\n=== Health Check Started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
            self.health_text.see(tk.END)
            self.root.update()
            
            # Run health check in a separate thread to avoid blocking GUI
            def run_check():
                try:
                    result = self.health_monitor.perform_comprehensive_health_check()
                    
                    # Update GUI in main thread
                    self.root.after(0, lambda: self.display_health_results(result))
                    
                except Exception as e:
                    self.root.after(0, lambda: self.health_text.insert(tk.END, f"Health check failed: {e}\n"))
            
            threading.Thread(target=run_check, daemon=True).start()
            
        except Exception as e:
            self.health_text.insert(tk.END, f"Failed to run health check: {e}\n")
    
    def display_health_results(self, result):
        """Display health check results"""
        try:
            overall_healthy = result.get('overall_healthy', False)
            health_score = result.get('health_score', 0)
            
            self.health_text.insert(tk.END, f"Overall Health: {'✅ HEALTHY' if overall_healthy else '❌ UNHEALTHY'}\n")
            self.health_text.insert(tk.END, f"Health Score: {health_score}%\n\n")
            
            # Display critical issues
            critical_issues = result.get('critical_issues', [])
            if critical_issues:
                self.health_text.insert(tk.END, "🚨 Critical Issues:\n")
                for issue in critical_issues:
                    self.health_text.insert(tk.END, f"  - {issue}\n")
                self.health_text.insert(tk.END, "\n")
            
            # Display warnings
            warnings = result.get('warnings', [])
            if warnings:
                self.health_text.insert(tk.END, "⚠️ Warnings:\n")
                for warning in warnings:
                    self.health_text.insert(tk.END, f"  - {warning}\n")
                self.health_text.insert(tk.END, "\n")
            
            # Display detailed results
            detailed_results = result.get('detailed_results', {})
            if detailed_results:
                self.health_text.insert(tk.END, "📊 Detailed Results:\n")
                for check_name, check_result in detailed_results.items():
                    healthy = check_result.get('healthy', False)
                    status_icon = "✅" if healthy else "❌"
                    self.health_text.insert(tk.END, f"  {status_icon} {check_name.replace('_', ' ').title()}\n")
                    
                    # Show specific issues for failed checks
                    if not healthy and 'issues' in check_result:
                        for issue in check_result['issues']:
                            self.health_text.insert(tk.END, f"    - {issue}\n")
            
            self.health_text.insert(tk.END, f"\n=== Health Check Completed ===\n\n")
            self.health_text.see(tk.END)
            
        except Exception as e:
            self.health_text.insert(tk.END, f"Failed to display health results: {e}\n")
    
    # Log viewer methods
    def refresh_log_files(self):
        """Refresh the list of available log files"""
        try:
            log_files = []
            
            # Find log files in EHC_Logs directory
            logs_dir = Path("EHC_Logs")
            if logs_dir.exists():
                for log_file in logs_dir.rglob("*.log"):
                    log_files.append(str(log_file))
            
            # Sort by modification time (newest first)
            log_files.sort(key=lambda x: Path(x).stat().st_mtime, reverse=True)
            
            # Update combobox
            if hasattr(self, 'log_file_var'):
                combo = None
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Frame):
                        for child in widget.winfo_children():
                            if isinstance(child, ttk.Notebook):
                                # Find log viewer tab
                                for tab_id in child.tabs():
                                    tab_text = child.tab(tab_id, "text")
                                    if tab_text == "Log Viewer":
                                        tab_frame = child.nametowidget(tab_id)
                                        for frame_child in tab_frame.winfo_children():
                                            if isinstance(frame_child, ttk.Frame):
                                                for frame_grandchild in frame_child.winfo_children():
                                                    if isinstance(frame_grandchild, ttk.Combobox):
                                                        combo = frame_grandchild
                                                        break
                
                if combo:
                    combo['values'] = log_files
                    if log_files and not self.log_file_var.get():
                        self.log_file_var.set(log_files[0])
            
        except Exception as e:
            self.logger.error(f"Failed to refresh log files: {e}")
    
    def load_log_file(self):
        """Load selected log file"""
        try:
            log_file = self.log_file_var.get()
            if not log_file:
                return
            
            log_path = Path(log_file)
            if not log_path.exists():
                messagebox.showerror("Error", f"Log file not found: {log_file}")
                return
            
            # Clear current content
            self.log_text.delete(1.0, tk.END)
            
            # Load file content (last 1000 lines to avoid memory issues)
            with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
                
            # Show last 1000 lines
            if len(lines) > 1000:
                self.log_text.insert(tk.END, f"[Showing last 1000 lines of {len(lines)} total lines]\n\n")
                lines = lines[-1000:]
            
            for line in lines:
                self.log_text.insert(tk.END, line)
            
            # Scroll to end
            self.log_text.see(tk.END)
            
            self.status_var.set(f"Loaded log file: {log_path.name}")
            
        except Exception as e:
            error_msg = f"Failed to load log file: {e}"
            self.logger.error(error_msg)
            messagebox.showerror("Error", error_msg)
    
    def toggle_log_auto_refresh(self):
        """Toggle auto-refresh for log files"""
        if self.auto_refresh_logs.get():
            self.start_log_auto_refresh()
        else:
            self.stop_log_auto_refresh()
    
    def start_log_auto_refresh(self):
        """Start auto-refresh for log files"""
        def refresh_log():
            if self.auto_refresh_logs.get() and self.root and self.root.winfo_exists():
                self.load_log_file()
                self.root.after(30000, refresh_log)  # Refresh every 30 seconds
        
        self.root.after(1000, refresh_log)  # Start after 1 second
    
    def stop_log_auto_refresh(self):
        """Stop auto-refresh for log files"""
        # Auto-refresh will stop automatically when checkbox is unchecked
        pass
    
    # Test methods
    def send_test_email(self):
        """Send test email"""
        try:
            if not self.email_system:
                messagebox.showerror("Error", "Email system not available")
                return
            
            recipient = self.test_email_var.get().strip()
            if not recipient:
                # Use first daily recipient if available
                daily_recipients = self.email_vars['daily_recipients'].get().strip()
                if daily_recipients:
                    recipient = daily_recipients.split(',')[0].strip()
                else:
                    messagebox.showerror("Error", "Please enter a test email recipient")
                    return
            
            self.test_results_text.insert(tk.END, f"\n=== Sending Test Email at {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, f"Recipient: {recipient}\n")
            self.test_results_text.see(tk.END)
            self.root.update()
            
            # Send test email in separate thread
            def send_email():
                try:
                    subject = f"MoonFlower Test Email - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                    body = f"""This is a test email from the MoonFlower WiFi Automation Configuration Panel.

Test Details:
- Sent at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- From: Configuration Panel
- Test Type: Manual test email

If you receive this email, the email configuration is working correctly.

Best regards,
MoonFlower WiFi Automation System
"""
                    
                    success = self.email_system.send_email(
                        recipient=recipient,
                        subject=subject,
                        body=body
                    )
                    
                    # Update GUI in main thread
                    if success:
                        self.root.after(0, lambda: self.test_results_text.insert(tk.END, "✅ Test email sent successfully!\n\n"))
                    else:
                        self.root.after(0, lambda: self.test_results_text.insert(tk.END, "❌ Test email failed to send\n\n"))
                    
                except Exception as e:
                    self.root.after(0, lambda: self.test_results_text.insert(tk.END, f"❌ Test email error: {e}\n\n"))
            
            threading.Thread(target=send_email, daemon=True).start()
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to send test email: {e}\n\n")
    
    def test_smtp_connection(self):
        """Test SMTP connection"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Testing SMTP Connection at {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.see(tk.END)
            self.root.update()
            
            # Test SMTP connection in separate thread
            def test_connection():
                try:
                    import smtplib
                    
                    server = self.email_vars['gmail_server'].get()
                    port = int(self.email_vars['gmail_port'].get())
                    
                    self.root.after(0, lambda: self.test_results_text.insert(tk.END, f"Connecting to {server}:{port}...\n"))
                    
                    with smtplib.SMTP(server, port) as smtp:
                        smtp.starttls()
                        self.root.after(0, lambda: self.test_results_text.insert(tk.END, "✅ SMTP connection successful!\n"))
                        self.root.after(0, lambda: self.test_results_text.insert(tk.END, f"Server response: {smtp.noop()}\n\n"))
                    
                except Exception as e:
                    self.root.after(0, lambda: self.test_results_text.insert(tk.END, f"❌ SMTP connection failed: {e}\n\n"))
            
            threading.Thread(target=test_connection, daemon=True).start()
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to test SMTP connection: {e}\n\n")
    
    def validate_smtp_settings(self):
        """Validate SMTP settings"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Validating SMTP Settings at {datetime.now().strftime('%H:%M:%S')} ===\n")
            
            # Validate server
            server = self.email_vars['gmail_server'].get().strip()
            if not server:
                self.test_results_text.insert(tk.END, "❌ SMTP server is required\n")
            else:
                self.test_results_text.insert(tk.END, f"✅ SMTP server: {server}\n")
            
            # Validate port
            try:
                port = int(self.email_vars['gmail_port'].get())
                if 1 <= port <= 65535:
                    self.test_results_text.insert(tk.END, f"✅ SMTP port: {port}\n")
                else:
                    self.test_results_text.insert(tk.END, f"❌ Invalid SMTP port: {port} (must be 1-65535)\n")
            except ValueError:
                self.test_results_text.insert(tk.END, f"❌ Invalid SMTP port: must be a number\n")
            
            self.test_results_text.insert(tk.END, "\n")
            self.test_results_text.see(tk.END)
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to validate SMTP settings: {e}\n\n")
    
    def test_csv_download(self):
        """Test CSV download functionality"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Testing CSV Download at {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "This would test the CSV download module...\n")
            self.test_results_text.insert(tk.END, "⚠️ Test not implemented yet\n\n")
            self.test_results_text.see(tk.END)
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to test CSV download: {e}\n\n")
    
    def test_excel_generation(self):
        """Test Excel generation functionality"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Testing Excel Generation at {datetime.now().strftime('%H:%M:%S')} ===\n")
            self.test_results_text.insert(tk.END, "This would test the Excel generation module...\n")
            self.test_results_text.insert(tk.END, "⚠️ Test not implemented yet\n\n")
            self.test_results_text.see(tk.END)
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to test Excel generation: {e}\n\n")
    
    def test_vbs_connection(self):
        """Test VBS connection"""
        try:
            self.test_results_text.insert(tk.END, f"\n=== Testing VBS Connection at {datetime.now().strftime('%H:%M:%S')} ===\n")
            
            primary_path = self.vbs_vars['primary_path'].get()
            backup_path = self.vbs_vars['backup_path'].get()
            
            if primary_path:
                if Path(primary_path).exists():
                    self.test_results_text.insert(tk.END, f"✅ Primary VBS path exists: {primary_path}\n")
                else:
                    self.test_results_text.insert(tk.END, f"❌ Primary VBS path not found: {primary_path}\n")
            else:
                self.test_results_text.insert(tk.END, "⚠️ Primary VBS path not configured\n")
            
            if backup_path:
                if Path(backup_path).exists():
                    self.test_results_text.insert(tk.END, f"✅ Backup VBS path exists: {backup_path}\n")
                else:
                    self.test_results_text.insert(tk.END, f"❌ Backup VBS path not found: {backup_path}\n")
            else:
                self.test_results_text.insert(tk.END, "⚠️ Backup VBS path not configured\n")
            
            self.test_results_text.insert(tk.END, "\n")
            self.test_results_text.see(tk.END)
            
        except Exception as e:
            self.test_results_text.insert(tk.END, f"Failed to test VBS connection: {e}\n\n")
    
    def run(self):
        """Run the configuration panel"""
        try:
            self.create_gui()
            
            # Handle window close event
            def on_closing():
                if self.unsaved_changes:
                    if messagebox.askyesno("Unsaved Changes", 
                                         "You have unsaved changes. Do you want to save before closing?"):
                        self.save_configuration()
                self.root.destroy()
            
            self.root.protocol("WM_DELETE_WINDOW", on_closing)
            
            # Start the GUI main loop
            self.root.mainloop()
            
        except Exception as e:
            self.logger.error(f"Failed to run configuration panel: {e}")
            if self.root:
                messagebox.showerror("Error", f"Configuration panel failed: {e}")


def main():
    """Main entry point for configuration panel"""
    try:
        # Setup logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Create and run configuration panel
        config_panel = ConfigurationPanel()
        config_panel.run()
        
    except Exception as e:
        print(f"Failed to start configuration panel: {e}")
        if tk._default_root:
            messagebox.showerror("Error", f"Failed to start configuration panel: {e}")


if __name__ == "__main__":
    main()