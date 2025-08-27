"""
Configuration Manager for MoonFlower WiFi Automation
Handles loading and validation of configuration settings
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
import logging

class ConfigManager:
    """Manages configuration settings for the automation system"""
    
    def __init__(self, config_path: str = "config/settings.json"):
        self.config_path = Path(config_path)
        self.config: Dict[str, Any] = {}
        self.logger = logging.getLogger(__name__)
        self.load_config()
    
    def load_config(self) -> None:
        """Load configuration from JSON file"""
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                self.logger.info(f"Configuration loaded from {self.config_path}")
            else:
                self.logger.warning(f"Configuration file not found: {self.config_path}")
                self.config = self._get_default_config()
                self.save_config()
        except Exception as e:
            self.logger.error(f"Failed to load configuration: {e}")
            self.config = self._get_default_config()
    
    def save_config(self) -> None:
        """Save current configuration to JSON file"""
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            self.logger.info(f"Configuration saved to {self.config_path}")
        except Exception as e:
            self.logger.error(f"Failed to save configuration: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key (supports dot notation)"""
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any) -> None:
        """Set configuration value by key (supports dot notation)"""
        keys = key.split('.')
        config = self.config
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value
        self.save_config()
    
    def get_ruckus_credentials(self) -> Dict[str, str]:
        """Get Ruckus controller credentials"""
        return {
            'url': self.get('ruckus_controller.url', ''),
            'username': self.get('ruckus_controller.username', ''),
            'password': self.get('ruckus_controller.password', '')
        }
    
    def get_networks(self, page: str = 'page_1') -> list:
        """Get network configurations for specified page"""
        return self.get(f'networks.{page}', [])
    
    def get_download_schedule(self) -> list:
        """Get download schedule times"""
        return self.get('download_schedule', ['09:30', '13:00'])
    
    def get_email_settings(self) -> Dict[str, Any]:
        """Get email configuration settings"""
        return self.get('email_settings', {})
    
    def get_system_settings(self) -> Dict[str, Any]:
        """Get system configuration settings"""
        return self.get('system_settings', {})
    
    def get_vbs_application_settings(self) -> Dict[str, Any]:
        """Get VBS application settings"""
        return self.get('vbs_application', {})
    
    def get_vbs_login_credentials(self) -> Dict[str, str]:
        """Get VBS application login credentials"""
        login_config = self.get('vbs_application.login', {})
        return {
            'username': login_config.get('username', ''),
            'password': login_config.get('password', ''),
            'database': login_config.get('database', '')
        }
    
    def get_vbs_application_paths(self) -> Dict[str, str]:
        """Get VBS application paths"""
        return {
            'primary_path': self.get('vbs_application.primary_path', ''),
            'backup_path': self.get('vbs_application.backup_path', '')
        }
    
    def get_vbs_timeout_settings(self) -> Dict[str, int]:
        """Get VBS timeout settings"""
        return self.get('vbs_application.timeout', {})
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default configuration"""
        return {
            "ruckus_controller": {
                "url": "https://51.38.163.73:8443/wsg/",
                "username": "admin",
                "password": "AdminFlower@123"
            },
            "networks": {
                "page_1": [
                    {
                        "name": "EHC TV",
                        "needs_clients_tab": True,
                        "selectors": {
                            "primary": "//span[contains(text(), 'EHC TV')]",
                            "fallback": "#ext-element-94"
                        }
                    },
                    {
                        "name": "EHC-15",
                        "needs_clients_tab": False,
                        "selectors": {
                            "primary": "//span[contains(text(), 'EHC-15')]",
                            "fallback": "#ext-element-93"
                        }
                    }
                ],
                "page_2": [
                    {
                        "name": "Reception Hall-Mobile",
                        "needs_clients_tab": True,
                        "selectors": {
                            "primary": "//span[contains(text(), 'Reception Hall-Mobile')]",
                            "fallback": "#ext-element-115"
                        }
                    },
                    {
                        "name": "Reception Hall-TV",
                        "needs_clients_tab": False,
                        "selectors": {
                            "primary": "//span[contains(text(), 'Reception Hall-TV')]",
                            "fallback": "#ext-element-114"
                        }
                    }
                ]
            },
            "download_schedule": ["09:30", "13:00"],
            "email_settings": {
                "send_weekdays_only": True,
                "send_all_days": False,
                "daily_recipients": [],
                "completion_recipients": [],
                "error_recipients": []
            },
            "system_settings": {
                "auto_restart": True,
                "debug_mode": False,
                "max_retry_attempts": 3,
                "retry_delay": 30
            },
            "vbs_application": {
                "primary_path": "C:\\Program Files\\AbsonsItERP\\AbsonsItERP.exe",
                "backup_path": "D:\\AbsonsItERP\\AbsonsItERP.exe",
                "window_title": "AbsonsItERP",
                "login": {
                    "username": "admin",
                    "password": "admin123",
                    "database": "MoonFlower"
                },
                "timeout": {
                    "app_launch": 30,
                    "login": 20,
                    "navigation": 15,
                    "data_import": 7200,
                    "pdf_generation": 300
                }
            }
        }