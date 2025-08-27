#!/usr/bin/env python3
"""
Configuration Loader for Computer Vision Services
Handles loading and validation of CV configuration settings
"""

import json
import os
import logging
from typing import Dict, Any, Optional
from pathlib import Path

class CVConfigLoader:
    """Configuration loader for computer vision services"""
    
    def __init__(self, config_path: str = "config/cv_config.json"):
        self.config_path = config_path
        self.config: Dict[str, Any] = {}
        self.logger = self._setup_logging()
        self.load_config()
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for config loader"""
        logger = logging.getLogger("CVConfigLoader")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def load_config(self) -> bool:
        """Load configuration from JSON file"""
        try:
            if not os.path.exists(self.config_path):
                self.logger.error(f"Configuration file not found: {self.config_path}")
                self._create_default_config()
                return False
            
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            
            self.logger.info(f"Configuration loaded successfully from {self.config_path}")
            return self._validate_config()
            
        except Exception as e:
            self.logger.error(f"Failed to load configuration: {e}")
            return False
    
    def _validate_config(self) -> bool:
        """Validate configuration structure and values"""
        try:
            required_sections = ['ocr_settings', 'template_matching', 'smart_automation', 'vbs_specific']
            
            for section in required_sections:
                if section not in self.config:
                    self.logger.error(f"Missing required configuration section: {section}")
                    return False
            
            # Validate OCR settings
            ocr_config = self.config['ocr_settings']
            if ocr_config['confidence_threshold'] < 0 or ocr_config['confidence_threshold'] > 1:
                self.logger.warning("OCR confidence threshold should be between 0 and 1")
            
            # Validate template matching settings
            template_config = self.config['template_matching']
            if template_config['confidence_threshold'] < 0 or template_config['confidence_threshold'] > 1:
                self.logger.warning("Template matching confidence threshold should be between 0 and 1")
            
            # Create template directory if it doesn't exist
            template_dir = template_config.get('template_directory', 'vbs/templates')
            os.makedirs(template_dir, exist_ok=True)
            
            # Create debug directory if needed
            if self.config.get('debugging', {}).get('save_debug_images', False):
                debug_dir = self.config['debugging'].get('debug_image_path', 'debug_images')
                os.makedirs(debug_dir, exist_ok=True)
            
            self.logger.info("Configuration validation completed successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Configuration validation failed: {e}")
            return False
    
    def _create_default_config(self):
        """Create default configuration file"""
        try:
            default_config = {
                "ocr_settings": {
                    "tesseract_path": "tesseract",
                    "language": "eng",
                    "confidence_threshold": 0.7,
                    "page_segmentation_mode": 6
                },
                "template_matching": {
                    "confidence_threshold": 0.8,
                    "template_directory": "vbs/templates"
                },
                "smart_automation": {
                    "method_priority": ["ocr", "template", "coordinates"],
                    "max_retries": 3
                },
                "vbs_specific": {
                    "window_title_patterns": ["absons", "arabian", "moonflower", "erp"]
                }
            }
            
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2)
            
            self.logger.info(f"Default configuration created at {self.config_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to create default configuration: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value by key (supports dot notation)"""
        try:
            keys = key.split('.')
            value = self.config
            
            for k in keys:
                if isinstance(value, dict) and k in value:
                    value = value[k]
                else:
                    return default
            
            return value
            
        except Exception:
            return default
    
    def get_ocr_config(self) -> Dict[str, Any]:
        """Get OCR configuration"""
        return self.config.get('ocr_settings', {})
    
    def get_template_config(self) -> Dict[str, Any]:
        """Get template matching configuration"""
        return self.config.get('template_matching', {})
    
    def get_automation_config(self) -> Dict[str, Any]:
        """Get smart automation configuration"""
        return self.config.get('smart_automation', {})
    
    def get_vbs_config(self) -> Dict[str, Any]:
        """Get VBS-specific configuration"""
        return self.config.get('vbs_specific', {})
    
    def get_performance_config(self) -> Dict[str, Any]:
        """Get performance configuration"""
        return self.config.get('performance', {})
    
    def get_debug_config(self) -> Dict[str, Any]:
        """Get debugging configuration"""
        return self.config.get('debugging', {})
    
    def update_config(self, key: str, value: Any) -> bool:
        """Update configuration value"""
        try:
            keys = key.split('.')
            config_ref = self.config
            
            # Navigate to the parent of the target key
            for k in keys[:-1]:
                if k not in config_ref:
                    config_ref[k] = {}
                config_ref = config_ref[k]
            
            # Set the value
            config_ref[keys[-1]] = value
            
            # Save to file
            return self.save_config()
            
        except Exception as e:
            self.logger.error(f"Failed to update configuration: {e}")
            return False
    
    def save_config(self) -> bool:
        """Save current configuration to file"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
            
            self.logger.info("Configuration saved successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save configuration: {e}")
            return False
    
    def reload_config(self) -> bool:
        """Reload configuration from file"""
        return self.load_config()

# Global configuration instance
_config_instance: Optional[CVConfigLoader] = None

def get_cv_config() -> CVConfigLoader:
    """Get global configuration instance"""
    global _config_instance
    if _config_instance is None:
        _config_instance = CVConfigLoader()
    return _config_instance

def reload_cv_config():
    """Reload global configuration"""
    global _config_instance
    if _config_instance:
        _config_instance.reload_config()
    else:
        _config_instance = CVConfigLoader()

if __name__ == "__main__":
    # Test configuration loading
    config = CVConfigLoader()
    print("OCR Config:", config.get_ocr_config())
    print("Template Config:", config.get_template_config())
    print("VBS Config:", config.get_vbs_config())