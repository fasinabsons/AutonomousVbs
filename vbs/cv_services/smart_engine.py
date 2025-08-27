#!/usr/bin/env python3
"""
Smart Automation Engine for VBS Phases
Orchestrates multiple automation methods with intelligent fallback and performance tracking
"""

import time
import logging
import os
import json
from typing import List, Dict, Tuple, Optional, Any, Callable
from dataclasses import dataclass, asdict
from enum import Enum
import cv2
import numpy as np
import pyautogui
import win32gui
import win32api
import win32con
from .ocr_service import OCRService, TextMatch, OCRResult
from .template_service import TemplateService, TemplateMatch, TemplateResult
from .config_loader import get_cv_config

class AutomationMethod(Enum):
    """Available automation methods in priority order"""
    OCR = "ocr"
    TEMPLATE = "template"
    COORDINATES = "coordinates"

@dataclass
class AutomationAction:
    """Represents an automation action to be performed"""
    action_type: str  # click, type, read, wait, etc.
    target_text: Optional[str] = None
    target_template: Optional[str] = None
    coordinates: Optional[Tuple[int, int]] = None
    input_text: Optional[str] = None
    region: Optional[Tuple[int, int, int, int]] = None
    timeout: float = 10.0
    confidence_threshold: float = 0.8
    retry_count: int = 3
    parameters: Optional[Dict[str, Any]] = None

@dataclass
class AutomationResult:
    """Result of an automation operation"""
    success: bool
    method_used: str
    execution_time: float
    confidence: float = 0.0
    location: Optional[Tuple[int, int]] = None
    error_message: Optional[str] = None
    screenshot_path: Optional[str] = None
    debug_info: Optional[Dict[str, Any]] = None

class SmartAutomationEngine:
    """Smart automation engine with multi-method fallback and performance tracking"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        
        # Initialize services
        self.ocr_service = OCRService()
        self.template_service = TemplateService()
        
        # Configuration
        self.smart_config = self.config.get('smart_automation', {})
        self.method_priority = [AutomationMethod(m) for m in self.smart_config.get('method_priority', ['ocr', 'template', 'coordinates'])]
        self.max_retries = self.smart_config.get('max_retries', 3)
        self.retry_delay = self.smart_config.get('retry_delay', 1.0)
        self.exponential_backoff = self.smart_config.get('exponential_backoff', True)
        
        # Performance tracking
        self.performance_stats = {
            'total_operations': 0,
            'successful_operations': 0,
            'method_success_rates': {method.value: 0 for method in AutomationMethod},
            'method_execution_times': {method.value: [] for method in AutomationMethod},
            'error_categories': {},
            'cache_hits': 0,
            'cache_misses': 0
        }
        
        # Location cache for successful operations
        self.location_cache = {}
        self.cache_duration = self.config.get('performance.cache_duration_seconds', 300)
        
        # Current window handle
        self.current_window = None
        
        self.logger.info("Smart Automation Engine initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for smart engine"""
        logger = logging.getLogger("SmartEngine")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler if debug mode is enabled
            if self.config.get('debugging.debug_mode', False):
                try:
                    log_file = "EHC_Logs/smart_engine.log"
                    os.makedirs(os.path.dirname(log_file), exist_ok=True)
                    file_handler = logging.FileHandler(log_file, encoding='utf-8')
                    file_handler.setFormatter(formatter)
                    logger.addHandler(file_handler)
                except Exception:
                    pass
        
        return logger
    
    def set_target_window(self, window_title_pattern: str = None) -> bool:
        """Set the target window for automation"""
        try:
            if window_title_pattern:
                # Find window by title pattern
                def enum_windows_callback(hwnd, windows):
                    if win32gui.IsWindowVisible(hwnd):
                        window_title = win32gui.GetWindowText(hwnd)
                        if window_title and window_title_pattern.lower() in window_title.lower():
                            windows.append((hwnd, window_title))
                    return True
                
                windows = []
                win32gui.EnumWindows(enum_windows_callback, windows)
                
                if windows:
                    self.current_window = windows[0][0]  # Use first match
                    self.logger.info(f"Target window set: {windows[0][1]}")
                    return True
                else:
                    self.logger.warning(f"No window found matching pattern: {window_title_pattern}")
                    return False
            else:
                # Use foreground window
                self.current_window = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(self.current_window)
                self.logger.info(f"Using foreground window: {window_title}")
                return True
                
        except Exception as e:
            self.logger.error(f"Failed to set target window: {e}")
            return False
    
    def capture_screenshot(self, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[np.ndarray]:
        """Capture screenshot of target window or region"""
        try:
            if self.current_window:
                # Capture from specific window
                screenshot = self.ocr_service.capture_screenshot_region(self.current_window, region)
            else:
                # Capture from screen
                if region:
                    x, y, w, h = region
                    screenshot = pyautogui.screenshot(region=(x, y, w, h))
                    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                else:
                    screenshot = pyautogui.screenshot()
                    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
            
            return screenshot
            
        except Exception as e:
            self.logger.error(f"Screenshot capture failed: {e}")
            return None
    
    def execute_action(self, action: AutomationAction) -> AutomationResult:
        """Execute automation action using smart method selection"""
        start_time = time.time()
        self.performance_stats['total_operations'] += 1
        
        # Check cache first for location-based actions
        cache_key = self._generate_cache_key(action)
        cached_location = self._get_cached_location(cache_key)
        
        if cached_location and action.action_type in ['click', 'double_click', 'right_click']:
            try:
                result = self._perform_action_at_location(cached_location, action)
                if result.success:
                    self.performance_stats['cache_hits'] += 1
                    result.method_used = "cache"
                    self.logger.info(f"Used cached location for action: {action.action_type}")
                    return result
            except Exception as e:
                self.logger.warning(f"Cached location failed: {e}")
        
        self.performance_stats['cache_misses'] += 1
        
        # Try each method in priority order
        last_error = None
        screenshot = None
        
        for method in self.method_priority:
            try:
                self.logger.info(f"Attempting {method.value} method for {action.action_type}")
                
                # Capture screenshot if not already done
                if screenshot is None:
                    screenshot = self.capture_screenshot(action.region)
                    if screenshot is None:
                        raise Exception("Failed to capture screenshot")
                
                # Execute using specific method
                if method == AutomationMethod.OCR:
                    result = self._execute_ocr_method(action, screenshot)
                elif method == AutomationMethod.TEMPLATE:
                    result = self._execute_template_method(action, screenshot)
                elif method == AutomationMethod.COORDINATES:
                    result = self._execute_coordinate_method(action)
                else:
                    continue
                
                # Update performance stats
                execution_time = time.time() - start_time
                result.execution_time = execution_time
                self.performance_stats['method_execution_times'][method.value].append(execution_time)
                
                if result.success:
                    self.performance_stats['successful_operations'] += 1
                    self.performance_stats['method_success_rates'][method.value] += 1
                    
                    # Cache successful location
                    if result.location and cache_key:
                        self._cache_location(cache_key, result.location)
                    
                    # Save debug screenshot if enabled
                    if self.config.get('debugging.screenshot_failed_operations', False):
                        result.screenshot_path = self._save_debug_screenshot(screenshot, f"success_{method.value}_{int(time.time())}.png")
                    
                    self.logger.info(f"Action successful using {method.value} method")
                    return result
                else:
                    last_error = result.error_message
                    self.logger.warning(f"{method.value} method failed: {result.error_message}")
                    
            except Exception as e:
                last_error = str(e)
                self.logger.error(f"{method.value} method error: {e}")
                continue
        
        # All methods failed
        execution_time = time.time() - start_time
        error_category = self._categorize_error(last_error)
        self.performance_stats['error_categories'][error_category] = self.performance_stats['error_categories'].get(error_category, 0) + 1
        
        # Save debug screenshot for failed operation
        screenshot_path = None
        if screenshot is not None and self.config.get('debugging.screenshot_failed_operations', False):
            screenshot_path = self._save_debug_screenshot(screenshot, f"failed_{action.action_type}_{int(time.time())}.png")
        
        return AutomationResult(
            success=False,
            method_used="none",
            execution_time=execution_time,
            error_message=f"All methods failed. Last error: {last_error}",
            screenshot_path=screenshot_path
        )    def
 _execute_ocr_method(self, action: AutomationAction, screenshot: np.ndarray) -> AutomationResult:
        """Execute action using OCR text detection"""
        try:
            if not action.target_text:
                return AutomationResult(
                    success=False,
                    method_used="ocr",
                    execution_time=0,
                    error_message="No target text specified for OCR method"
                )
            
            # Find text using OCR
            matches = self.ocr_service.find_text(screenshot, action.target_text, case_sensitive=False)
            
            if not matches:
                return AutomationResult(
                    success=False,
                    method_used="ocr",
                    execution_time=0,
                    error_message=f"Text '{action.target_text}' not found"
                )
            
            # Use the match with highest confidence
            best_match = max(matches, key=lambda m: m.confidence)
            
            if best_match.confidence < action.confidence_threshold:
                return AutomationResult(
                    success=False,
                    method_used="ocr",
                    execution_time=0,
                    confidence=best_match.confidence,
                    error_message=f"OCR confidence {best_match.confidence:.2f} below threshold {action.confidence_threshold}"
                )
            
            # Perform action at detected location
            return self._perform_action_at_location(best_match.center, action, best_match.confidence)
            
        except Exception as e:
            return AutomationResult(
                success=False,
                method_used="ocr",
                execution_time=0,
                error_message=f"OCR method error: {str(e)}"
            )
    
    def _execute_template_method(self, action: AutomationAction, screenshot: np.ndarray) -> AutomationResult:
        """Execute action using template matching"""
        try:
            if not action.target_template:
                return AutomationResult(
                    success=False,
                    method_used="template",
                    execution_time=0,
                    error_message="No target template specified for template method"
                )
            
            # Find template
            result = self.template_service.find_template(screenshot, action.target_template)
            
            if not result.success or not result.matches:
                return AutomationResult(
                    success=False,
                    method_used="template",
                    execution_time=0,
                    error_message=f"Template '{action.target_template}' not found"
                )
            
            # Use the match with highest confidence
            best_match = max(result.matches, key=lambda m: m.confidence)
            
            if best_match.confidence < action.confidence_threshold:
                return AutomationResult(
                    success=False,
                    method_used="template",
                    execution_time=0,
                    confidence=best_match.confidence,
                    error_message=f"Template confidence {best_match.confidence:.2f} below threshold {action.confidence_threshold}"
                )
            
            # Calculate center of template match
            center = (best_match.location[0] + best_match.size[0] // 2,
                     best_match.location[1] + best_match.size[1] // 2)
            
            # Perform action at detected location
            return self._perform_action_at_location(center, action, best_match.confidence)
            
        except Exception as e:
            return AutomationResult(
                success=False,
                method_used="template",
                execution_time=0,
                error_message=f"Template method error: {str(e)}"
            )
    
    def _execute_coordinate_method(self, action: AutomationAction) -> AutomationResult:
        """Execute action using hardcoded coordinates (fallback)"""
        try:
            if not action.coordinates:
                return AutomationResult(
                    success=False,
                    method_used="coordinates",
                    execution_time=0,
                    error_message="No coordinates specified for coordinate method"
                )
            
            # Perform action at specified coordinates
            return self._perform_action_at_location(action.coordinates, action, 1.0)
            
        except Exception as e:
            return AutomationResult(
                success=False,
                method_used="coordinates",
                execution_time=0,
                error_message=f"Coordinate method error: {str(e)}"
            )
    
    def _perform_action_at_location(self, location: Tuple[int, int], action: AutomationAction, confidence: float = 1.0) -> AutomationResult:
        """Perform the actual automation action at the specified location"""
        try:
            x, y = location
            
            # Add small delay for UI responsiveness
            ui_delay = self.config.get('vbs_specific.ui_response_delay', 0.5)
            time.sleep(ui_delay)
            
            if action.action_type == "click":
                pyautogui.click(x, y)
                
            elif action.action_type == "double_click":
                pyautogui.doubleClick(x, y)
                
            elif action.action_type == "right_click":
                pyautogui.rightClick(x, y)
                
            elif action.action_type == "type":
                if action.input_text:
                    pyautogui.click(x, y)  # Click first to focus
                    time.sleep(0.2)
                    pyautogui.typewrite(action.input_text)
                else:
                    raise ValueError("No input text provided for type action")
                    
            elif action.action_type == "read":
                # For read actions, we just return the location
                pass
                
            elif action.action_type == "wait":
                # Wait for element to appear (already found if we're here)
                pass
                
            else:
                raise ValueError(f"Unsupported action type: {action.action_type}")
            
            return AutomationResult(
                success=True,
                method_used="",  # Will be set by caller
                execution_time=0,  # Will be set by caller
                confidence=confidence,
                location=location
            )
            
        except Exception as e:
            return AutomationResult(
                success=False,
                method_used="",
                execution_time=0,
                error_message=f"Action execution failed: {str(e)}"
            )
    
    def _generate_cache_key(self, action: AutomationAction) -> Optional[str]:
        """Generate cache key for action"""
        try:
            if action.target_text:
                return f"text_{action.target_text}_{action.action_type}"
            elif action.target_template:
                return f"template_{action.target_template}_{action.action_type}"
            else:
                return None
        except Exception:
            return None
    
    def _get_cached_location(self, cache_key: str) -> Optional[Tuple[int, int]]:
        """Get cached location if still valid"""
        if not cache_key or cache_key not in self.location_cache:
            return None
        
        cached_data = self.location_cache[cache_key]
        if time.time() - cached_data['timestamp'] > self.cache_duration:
            # Cache expired
            del self.location_cache[cache_key]
            return None
        
        return cached_data['location']
    
    def _cache_location(self, cache_key: str, location: Tuple[int, int]):
        """Cache successful location"""
        if cache_key:
            self.location_cache[cache_key] = {
                'location': location,
                'timestamp': time.time()
            }
    
    def _categorize_error(self, error_message: str) -> str:
        """Categorize error for analytics"""
        if not error_message:
            return "unknown"
        
        error_lower = error_message.lower()
        
        if "not found" in error_lower:
            return "element_not_found"
        elif "confidence" in error_lower or "threshold" in error_lower:
            return "low_confidence"
        elif "screenshot" in error_lower or "capture" in error_lower:
            return "screenshot_failure"
        elif "timeout" in error_lower:
            return "timeout"
        elif "coordinate" in error_lower:
            return "coordinate_error"
        elif "ocr" in error_lower:
            return "ocr_error"
        elif "template" in error_lower:
            return "template_error"
        else:
            return "other"
    
    def _save_debug_screenshot(self, screenshot: np.ndarray, filename: str) -> Optional[str]:
        """Save debug screenshot"""
        try:
            debug_dir = self.config.get('debugging.debug_image_path', 'debug_images')
            os.makedirs(debug_dir, exist_ok=True)
            
            debug_path = os.path.join(debug_dir, filename)
            cv2.imwrite(debug_path, screenshot)
            
            return debug_path
            
        except Exception as e:
            self.logger.error(f"Failed to save debug screenshot: {e}")
            return None
    
    def execute_vbs_phase_action(self, phase_name: str, action_name: str, **kwargs) -> AutomationResult:
        """Execute a predefined VBS phase action"""
        try:
            # Load VBS-specific action definitions
            vbs_actions = self._get_vbs_actions()
            
            if phase_name not in vbs_actions:
                return AutomationResult(
                    success=False,
                    method_used="none",
                    execution_time=0,
                    error_message=f"Unknown VBS phase: {phase_name}"
                )
            
            if action_name not in vbs_actions[phase_name]:
                return AutomationResult(
                    success=False,
                    method_used="none",
                    execution_time=0,
                    error_message=f"Unknown action '{action_name}' in phase '{phase_name}'"
                )
            
            # Get action definition and merge with kwargs
            action_def = vbs_actions[phase_name][action_name].copy()
            action_def.update(kwargs)
            
            # Create AutomationAction object
            action = AutomationAction(**action_def)
            
            # Execute with retry logic
            return self.execute_action_with_retry(action)
            
        except Exception as e:
            return AutomationResult(
                success=False,
                method_used="none",
                execution_time=0,
                error_message=f"VBS phase action error: {str(e)}"
            )
    
    def execute_action_with_retry(self, action: AutomationAction) -> AutomationResult:
        """Execute action with retry logic"""
        last_result = None
        
        for attempt in range(self.max_retries):
            if attempt > 0:
                delay = self.retry_delay
                if self.exponential_backoff:
                    delay *= (2 ** (attempt - 1))
                
                self.logger.info(f"Retrying action (attempt {attempt + 1}/{self.max_retries}) after {delay}s delay")
                time.sleep(delay)
            
            result = self.execute_action(action)
            
            if result.success:
                return result
            
            last_result = result
            self.logger.warning(f"Attempt {attempt + 1} failed: {result.error_message}")
        
        # All retries failed
        return last_result or AutomationResult(
            success=False,
            method_used="none",
            execution_time=0,
            error_message="All retry attempts failed"
        )
    
    def _get_vbs_actions(self) -> Dict[str, Dict[str, Dict[str, Any]]]:
        """Get VBS-specific action definitions"""
        return {
            "phase2_navigation": {
                "click_sales_distribution": {
                    "action_type": "click",
                    "target_text": "Sales & Distribution",
                    "target_template": "sales_distribution_menu",
                    "confidence_threshold": 0.7
                },
                "click_pos": {
                    "action_type": "click",
                    "target_text": "POS",
                    "target_template": "pos_menu",
                    "confidence_threshold": 0.7
                },
                "click_wifi_registration": {
                    "action_type": "click",
                    "target_text": "WiFi User Registration",
                    "target_template": "wifi_registration_menu",
                    "confidence_threshold": 0.7
                }
            },
            "phase3_upload": {
                "click_import_button": {
                    "action_type": "click",
                    "target_text": "Import",
                    "target_template": "import_button",
                    "confidence_threshold": 0.8
                },
                "click_browse_button": {
                    "action_type": "click",
                    "target_text": "Browse",
                    "target_template": "browse_button",
                    "confidence_threshold": 0.8
                },
                "click_update_button": {
                    "action_type": "click",
                    "target_text": "Update",
                    "target_template": "update_button",
                    "confidence_threshold": 0.8
                }
            },
            "phase4_report": {
                "click_reports_menu": {
                    "action_type": "click",
                    "target_text": "Reports",
                    "target_template": "reports_menu",
                    "confidence_threshold": 0.7
                },
                "click_print_button": {
                    "action_type": "click",
                    "target_text": "Print",
                    "target_template": "print_button",
                    "confidence_threshold": 0.8
                },
                "click_export_button": {
                    "action_type": "click",
                    "target_text": "Export",
                    "target_template": "export_button",
                    "confidence_threshold": 0.8
                }
            }
        }
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get comprehensive performance statistics"""
        stats = self.performance_stats.copy()
        
        # Calculate success rate
        if stats['total_operations'] > 0:
            stats['overall_success_rate'] = stats['successful_operations'] / stats['total_operations']
        else:
            stats['overall_success_rate'] = 0.0
        
        # Calculate average execution times per method
        stats['average_execution_times'] = {}
        for method, times in stats['method_execution_times'].items():
            if times:
                stats['average_execution_times'][method] = sum(times) / len(times)
            else:
                stats['average_execution_times'][method] = 0.0
        
        # Add OCR and template service stats
        stats['ocr_stats'] = self.ocr_service.get_performance_stats()
        stats['template_stats'] = self.template_service.get_performance_stats()
        
        return stats
    
    def reset_performance_stats(self):
        """Reset all performance statistics"""
        self.performance_stats = {
            'total_operations': 0,
            'successful_operations': 0,
            'method_success_rates': {method.value: 0 for method in AutomationMethod},
            'method_execution_times': {method.value: [] for method in AutomationMethod},
            'error_categories': {},
            'cache_hits': 0,
            'cache_misses': 0
        }
        
        # Reset service stats
        self.ocr_service.reset_stats()
        self.template_service.reset_stats()
        
        # Clear cache
        self.location_cache.clear()
    
    def save_performance_report(self, filepath: str = None):
        """Save performance report to file"""
        try:
            if not filepath:
                timestamp = int(time.time())
                filepath = f"EHC_Logs/smart_engine_performance_{timestamp}.json"
            
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            report = {
                'timestamp': time.time(),
                'performance_stats': self.get_performance_stats(),
                'configuration': {
                    'method_priority': [m.value for m in self.method_priority],
                    'max_retries': self.max_retries,
                    'retry_delay': self.retry_delay,
                    'exponential_backoff': self.exponential_backoff
                }
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Performance report saved to: {filepath}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save performance report: {e}")
            return False

if __name__ == "__main__":
    # Test smart automation engine
    engine = SmartAutomationEngine()
    print("Smart Automation Engine initialized successfully")
    
    # Test action
    test_action = AutomationAction(
        action_type="click",
        target_text="Test Button",
        confidence_threshold=0.7
    )
    
    result = engine.execute_action(test_action)
    print(f"Test result: {result}")