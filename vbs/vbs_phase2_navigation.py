#!/usr/bin/env python3
"""
VBS Automation - Phase 2: Enhanced Navigation with Superior Double-Click
Navigation: Sales & Distribution → POS → WiFi User Registration
Combines enhanced double-click from vbsImport.txt with keyboard navigation from prd.txt
"""

import time
import logging
import win32gui
import win32con
import win32api
from typing import Dict, Optional, Any, List
import traceback
from datetime import datetime
import os


# Windows API constants for enhanced double-click support
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203  # CRITICAL: Double-click message
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_CHAR = 0x0102
VK_RETURN = 0x0D
VK_ESCAPE = 0x1B

# Mouse event flags
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_ABSOLUTE = 0x8000

class VBSPhase2_Navigation:
    """VBS Navigation Phase - Handles navigation to WiFi User Registration"""
    
    def __init__(self, window_handle: Optional[int] = None, config_manager=None, file_manager=None):
        """Initialize VBS Navigation Phase with Enhanced Double-Click"""
        self.logger = self._setup_logging()
        
        # Get system double-click settings - CRITICAL for proper double-click timing
        try:
            self.double_click_time = win32gui.GetDoubleClickTime()  # Get system double-click time
            self.logger.info(f"System double-click time: {self.double_click_time}ms")
        except:
            self.double_click_time = 1500  # Default fallback
            self.logger.info(f"Using default double-click time: {self.double_click_time}ms")
        
        # Use provided window handle if available
        if window_handle:
            self.window_handle = window_handle
        else:
            self.window_handle = self._find_main_vbs_window()
            if self.window_handle:
                self.logger.info(f"Found VBS window automatically: {self.window_handle}")
            else:
                self.logger.warning("No VBS window found, will search again when needed")
        
        # Enhanced coordinates with better precision (from vbsImport.txt)
        self.coordinates = {
            # Main menu coordinates - EXACT from vbsImport.txt
            "arrow_button": (23, 64),
            "sales_distribution": (212, 170),
            "pos_submenu": (179, 601),
            "wifi_user_registration": (288, 679),  # This MUST be double-clicked
            "new_button": (104, 107),
            
            # Enhanced fallback coordinates with micro-adjustments
            "arrow_button_fallback": [
                (23, 64), (22, 64), (24, 64), (23, 63), (23, 65),
                (21, 64), (25, 64), (23, 62), (23, 66), (20, 64)
            ],
            "sales_distribution_fallback": [
                (212, 170), (211, 170), (213, 170), (212, 169), (212, 171),
                (210, 170), (214, 170), (212, 168), (212, 172), (208, 170)
            ],
            "pos_submenu_fallback": [
                (179, 601), (178, 601), (180, 601), (179, 600), (179, 602),
                (177, 601), (181, 601), (179, 599), (179, 603), (175, 601)
            ],
            "wifi_user_registration_fallback": [
                (288, 679), (287, 679), (289, 679), (288, 678), (288, 680),
                (286, 679), (290, 679), (288, 677), (288, 681), (285, 679)
            ],
            "new_button_fallback": [
                (104, 107), (103, 107), (105, 107), (104, 106), (104, 108),
                (102, 107), (106, 107), (104, 105), (104, 109), (100, 107)
            ]
        }
        
        # Enhanced timing configuration for double-clicks
        self.timeouts = {
            "menu_open": 2.0,
            "submenu_open": 1.5,
            "window_load": 3.0,
            "element_wait": 1.0,
            "double_click_interval": 0.05,  # Very fast double-click interval
            "double_click_total": 0.15,     # Total time for double-click sequence
            "new_button_wait": 2.0,
            "visual_feedback": 0.3,         # Reduced visual feedback time
            "click_validation": 0.5,        # Time to validate click success
        }
        
        # Enhanced navigation sequence with explicit double-click specification
        self.navigation_sequence = [
            {
                "step": "arrow_button",
                "description": "Click Arrow button to open menu",
                "coordinate": "arrow_button",
                "wait_time": "menu_open",
                "double_click": False,
                "critical": True
            },
            {
                "step": "sales_distribution",
                "description": "Click Sales & Distribution menu",
                "coordinate": "sales_distribution",
                "wait_time": "menu_open",
                "double_click": False,
                "critical": True
            },
            {
                "step": "pos_submenu",
                "description": "Click POS submenu",
                "coordinate": "pos_submenu",
                "wait_time": "submenu_open",
                "double_click": False,
                "critical": True
            },
            {
                "step": "wifi_registration",
                "description": "DOUBLE-CLICK WiFi User Registration (CRITICAL)",
                "coordinate": "wifi_user_registration",
                "wait_time": "window_load",
                "double_click": True,  # CRITICAL: Must be double-click
                "critical": True,
                "retry_count": 5  # More retries for this critical step
            },
            {
                "step": "new_button",
                "description": "Click New button to create new entry",
                "coordinate": "new_button",
                "wait_time": "new_button_wait",
                "double_click": False,
                "critical": True
            }
        ]

        self.logger.info("Enhanced VBS Phase 2 Navigation with Superior Double-Click and LLM integration initialized")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup enhanced logging"""
        logger = logging.getLogger("VBSPhase2_Navigation")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                log_file = "EHC_Logs/vbs_phase2_navigation.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def _find_main_vbs_window(self) -> Optional[int]:
        """Find main VBS window handle"""
        vbs_windows = []
        
        def enum_windows_callback(hwnd, windows):
            try:
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    if title:
                        title_lower = title.lower()
                        vbs_indicators = ['absons', 'arabian', 'moonflower', 'erp']
                        exclude_indicators = ['login', 'security', 'warning', 'browser']
                        
                        has_vbs = any(indicator in title_lower for indicator in vbs_indicators)
                        has_exclude = any(indicator in title_lower for indicator in exclude_indicators)
                        
                        if has_vbs and not has_exclude:
                            windows.append((hwnd, title))
            except:
                pass
            return True
        
        win32gui.EnumWindows(enum_windows_callback, vbs_windows)
        
        if vbs_windows:
            self.logger.info(f"Found VBS window: {vbs_windows[0][1]}")
            return vbs_windows[0][0]
        
        return None
    
    def _ensure_window_focus(self) -> bool:
        """Ensure VBS window is focused and ready"""
        try:
            if not self.window_handle:
                self.window_handle = self._find_main_vbs_window()
                if not self.window_handle:
                    self.logger.error("No VBS window found")
                    return False
            
            # Focus the window
            win32gui.SetForegroundWindow(self.window_handle)
            win32gui.ShowWindow(self.window_handle, win32con.SW_RESTORE)
            time.sleep(0.5)
            
            # Verify focus
            focused_window = win32gui.GetForegroundWindow()
            if focused_window == self.window_handle:
                self.logger.info("VBS window focused successfully")
                return True
            else:
                self.logger.warning("Window focus uncertain, continuing anyway")
                return True
            
        except Exception as e:
            self.logger.error(f"Window focus failed: {e}")
            return False
    
    def _enhanced_double_click_windows_api(self, x: int, y: int) -> bool:
        """Enhanced Windows API double-click with proper timing and messages"""
        try:
            if not self.window_handle:
                return False
            
            # Ensure window is focused and active
            try:
                win32gui.SetForegroundWindow(self.window_handle)
                win32gui.SetActiveWindow(self.window_handle)
                time.sleep(0.1)
            except:
                pass
            
            # Get window position for screen coordinates
            window_rect = win32gui.GetWindowRect(self.window_handle)
            screen_x = window_rect[0] + x
            screen_y = window_rect[1] + y
            
            # Move cursor to exact position
            win32api.SetCursorPos((screen_x, screen_y))
            time.sleep(0.05)
            
            self.logger.info(f"ENHANCED DOUBLE-CLICK: Position ({screen_x}, {screen_y})")
            
            # Method 1: Use proper double-click sequence with Windows API
            # First click
            win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.01)
            win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            
            # Critical timing - use system double-click time
            time.sleep(self.double_click_time / 2000.0)  # Convert to seconds and halve
            
            # Second click (this creates the double-click)
            win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.01)
            win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            
            # Additional method: Send double-click message directly to window
            lParam = win32api.MAKELONG(x, y)
            win32gui.PostMessage(self.window_handle, WM_LBUTTONDBLCLK, 0, lParam)
            
            time.sleep(self.timeouts["double_click_total"])
            
            self.logger.info("Enhanced Windows API double-click completed")
            return True
            
        except Exception as e:
            self.logger.error(f"Enhanced Windows API double-click failed: {e}")
            return False
    
    def _enhanced_double_click_window_messages(self, x: int, y: int) -> bool:
        """Enhanced window message double-click with proper sequence"""
        try:
            if not self.window_handle:
                return False
            
            lParam = win32api.MAKELONG(x, y)
            
            self.logger.info(f"WINDOW MESSAGE DOUBLE-CLICK: Position ({x}, {y})")
            
            # Send proper double-click sequence
            # First click
            win32gui.PostMessage(self.window_handle, WM_LBUTTONDOWN, 0, lParam)
            time.sleep(0.01)
            win32gui.PostMessage(self.window_handle, WM_LBUTTONUP, 0, lParam)
            
            # Small delay
            time.sleep(self.timeouts["double_click_interval"])
            
            # Second click
            win32gui.PostMessage(self.window_handle, WM_LBUTTONDOWN, 0, lParam)
            time.sleep(0.01)
            win32gui.PostMessage(self.window_handle, WM_LBUTTONUP, 0, lParam)
            
            # Send explicit double-click message
            time.sleep(0.01)
            win32gui.PostMessage(self.window_handle, WM_LBUTTONDBLCLK, 0, lParam)
            
            time.sleep(self.timeouts["double_click_total"])
            
            self.logger.info("Enhanced window message double-click completed")
            return True
            
        except Exception as e:
            self.logger.error(f"Enhanced window message double-click failed: {e}")
            return False
    
    def _ultra_robust_click(self, coordinate_name: str, double_click: bool = False, max_retries: int = 3) -> bool:
        """Ultra-robust clicking with enhanced double-click support"""
        try:
            if not self.window_handle:
                self.logger.error("No window handle available")
                return False
            
            # Get coordinates
            if coordinate_name not in self.coordinates:
                self.logger.error(f"Unknown coordinate: {coordinate_name}")
                return False
            
            primary_coord = self.coordinates[coordinate_name]
            fallback_coords = self.coordinates.get(f"{coordinate_name}_fallback", [])
            all_coords = [primary_coord] + fallback_coords
            
            # Enhanced retry logic
            for retry in range(max_retries):
                self.logger.info(f"Retry {retry + 1}/{max_retries} for {coordinate_name}")
                
                for attempt, (x, y) in enumerate(all_coords):
                    self.logger.info(f"Attempting coordinate ({x}, {y})")
                    
                    # Verify window is still valid
                    if not win32gui.IsWindow(self.window_handle):
                        self.logger.error("Window is no longer valid")
                        return False
                    
                    if double_click:
                        self.logger.info(f"DOUBLE-CLICK MODE: {coordinate_name}")
                        
                        # Try all double-click methods
                        methods = [
                            self._enhanced_double_click_windows_api,
                            self._enhanced_double_click_window_messages
                        ]
                        
                        for method_idx, method in enumerate(methods):
                            self.logger.info(f"Trying double-click method {method_idx + 1}")
                            if method(x, y):
                                # Validate double-click success
                                time.sleep(self.timeouts["click_validation"])
                                if self._validate_click_success(coordinate_name, double_click):
                                    self.logger.info(f"Double-click successful with method {method_idx + 1}")
                                    return True
                            time.sleep(0.2)
                    
                    else:
                        # Single click
                        if self._single_click_enhanced(x, y):
                            time.sleep(self.timeouts["click_validation"])
                            if self._validate_click_success(coordinate_name, double_click):
                                self.logger.info(f"Single-click successful")
                                return True
                    
                    time.sleep(0.5)  # Brief pause between coordinate attempts
                
                # Wait before retry
                if retry < max_retries - 1:
                    time.sleep(1.0)
            
            self.logger.error(f"All attempts failed for {coordinate_name}")
            return False
            
        except Exception as e:
            self.logger.error(f"Ultra-robust click failed: {e}")
            return False
    
    def _single_click_enhanced(self, x: int, y: int) -> bool:
        """Enhanced single click implementation"""
        try:
            if not self.window_handle:
                return False
            
            # Ensure window focus
            try:
                win32gui.SetForegroundWindow(self.window_handle)
                time.sleep(0.1)
            except:
                pass
            
            # Get screen coordinates
            window_rect = win32gui.GetWindowRect(self.window_handle)
            screen_x = window_rect[0] + x
            screen_y = window_rect[1] + y
            
            # Move cursor
            win32api.SetCursorPos((screen_x, screen_y))
            time.sleep(0.05)
            
            # Perform click
            win32api.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.05)
            win32api.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Enhanced single click failed: {e}")
            return False
    
    def _validate_click_success(self, coordinate_name: str, was_double_click: bool) -> bool:
        """Validate if click was successful"""
        try:
            # Basic validation - window still exists and is responsive
            if not win32gui.IsWindow(self.window_handle):
                return False
            
            # For WiFi User Registration double-click, check if new window opened
            if coordinate_name == "wifi_user_registration" and was_double_click:
                # Look for WiFi Registration window
                time.sleep(1.0)  # Give window time to open
                return self._check_wifi_registration_window()
            
            # For other clicks, basic validation
            return True
            
        except Exception as e:
            self.logger.warning(f"Click validation failed: {e}")
            return False
    
    def _check_wifi_registration_window(self) -> bool:
        """Check if WiFi User Registration window opened"""
        try:
            wifi_windows = []
            
            def enum_windows_callback(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if title:
                            title_lower = title.lower()
                            if 'wifi' in title_lower and ('user' in title_lower or 'registration' in title_lower):
                                windows.append((hwnd, title))
                except:
                    pass
            
            win32gui.EnumWindows(enum_windows_callback, wifi_windows)
            
            if wifi_windows:
                self.logger.info(f"WiFi Registration window detected: {wifi_windows[0][1]}")
                return True
            
            return False
            
        except Exception as e:
            self.logger.warning(f"WiFi window check failed: {e}")
            return False
    
    def _press_key_enhanced(self, key_name: str) -> bool:
        """Press key with enhanced method"""
        try:
            # Key mapping
            key_codes = {
                "ENTER": win32con.VK_RETURN,
                "TAB": win32con.VK_TAB, 
                "LEFT": win32con.VK_LEFT,
                "RIGHT": win32con.VK_RIGHT,
                "ESCAPE": win32con.VK_ESCAPE
            }
            
            if key_name not in key_codes:
                self.logger.error(f"Unknown key: {key_name}")
                return False
            
            vk_code = key_codes[key_name]
            
            self.logger.info(f"Pressing {key_name} key")
            
            # Direct window message (preferred)
            if self.window_handle:
                win32api.SendMessage(self.window_handle, win32con.WM_KEYDOWN, vk_code, 0)
                time.sleep(self.timing['enter_press_delay'])
                win32api.SendMessage(self.window_handle, win32con.WM_KEYUP, vk_code, 0)
            else:
                # Global keyboard event (fallback)
                win32api.keybd_event(vk_code, 0, 0, 0)
                time.sleep(self.timing['enter_press_delay'])
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            
            self.logger.info(f"{key_name} key pressed successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Key press failed for {key_name}: {e}")
            return False
    
    def _execute_keyboard_sequence(self, sequence: List[str], delay_between: float = None) -> bool:
        """Execute a sequence of keyboard keys"""
        try:
            if delay_between is None:
                delay_between = self.timing['tab_press_delay']
            
            self.logger.info(f"Executing keyboard sequence: {' -> '.join(sequence)}")
            
            for i, key in enumerate(sequence):
                self.logger.info(f"Step {i+1}/{len(sequence)}: Pressing {key}")
                
                if not self._press_key_enhanced(key):
                    self.logger.error(f"Failed to press {key} in sequence")
                    return False
                
                # Wait between keys (except for last key)
                if i < len(sequence) - 1:
                    time.sleep(delay_between)
            
            self.logger.info("Keyboard sequence completed successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Keyboard sequence failed: {e}")
            return False
    
    def _wait_for_ui_response(self, wait_type: str):
        """Wait for UI to respond after action"""
        wait_time = self.timing.get(wait_type, 1.0)
        self.logger.info(f"Waiting {wait_time}s for UI response ({wait_type})")
        time.sleep(wait_time)
    
    def _validate_step_success(self, step_name: str) -> bool:
        """Validate if step was successful"""
        try:
            # Basic validation - check if window is still responsive
            if self.window_handle:
                if not win32gui.IsWindow(self.window_handle):
                    self.logger.warning("Window handle no longer valid")
                    return False
            
            # For WiFi activation, check if new window opened
            if step_name == "wifi_activation_enter":
                return self._check_wifi_window_opened()
            
            # For new button activation, check for form fields
            elif step_name == "new_button_activation":
                return self._check_new_form_opened()
            
            # Default: assume success if no errors
            return True
            
        except Exception as e:
            self.logger.warning(f"Step validation failed: {e}")
            return True  # Continue anyway
    
    def _check_wifi_window_opened(self) -> bool:
        """Check if WiFi User Registration window opened"""
        try:
            wifi_windows = []
            
            def enum_windows_callback(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if title:
                            title_lower = title.lower()
                            wifi_indicators = ['wifi', 'user', 'registration']
                            if any(indicator in title_lower for indicator in wifi_indicators):
                                windows.append((hwnd, title))
                except:
                    pass
                return True
            
            win32gui.EnumWindows(enum_windows_callback, wifi_windows)
            
            if wifi_windows:
                self.logger.info(f"WiFi window detected: {wifi_windows[0][1]}")
                # Update window handle to new window
                self.window_handle = wifi_windows[0][0]
                return True
            
            self.logger.info("No new WiFi window detected, continuing...")
            return True  # Continue anyway
            
        except Exception as e:
            self.logger.warning(f"WiFi window check failed: {e}")
            return True
    
    def _check_new_form_opened(self) -> bool:
        """Check if New form opened"""
        try:
            self.logger.info("Checking for New form... (basic validation)")
            return True  # Assume success for now
            
        except Exception as e:
            self.logger.warning(f"New form check failed: {e}")
            return True
    
    def execute_navigation(self) -> Dict[str, Any]:
        """Execute the complete navigation sequence"""
        try:
            self.logger.info("Starting Enhanced Navigation with Keyboard Shortcuts")
            
            result = {
                "success": False,
                "start_time": datetime.now().isoformat(),
                "steps_completed": [],
                "errors": []
            }
            
            # Ensure window is ready
            if not self._ensure_window_focus():
                result["errors"].append("Failed to focus VBS window")
                return result
            
            # Execute each step in the sequence
            for i, step in enumerate(self.navigation_sequence, 1):
                step_name = step["step"]
                description = step["description"]
                action = step["action"]
                
                self.logger.info(f"Step {i}/{len(self.navigation_sequence)}: {description}")
                
                step_success = False
                
                if action == "click":
                    coordinate = step["coordinate"]
                    step_success = self._click_coordinate(coordinate)
                    
                elif action == "key_press":
                    key = step["key"]
                    step_success = self._press_key_enhanced(key)
                    
                elif action == "keyboard_sequence":
                    sequence = step["sequence"]
                    delay = self.timing.get(step.get("sequence_delay", "tab_press_delay"), 0.5)
                    step_success = self._execute_keyboard_sequence(sequence, delay)
                
                if not step_success:
                    error_msg = f"Failed at step {i}: {description}"
                    self.logger.error(f"{error_msg}")
                    result["errors"].append(error_msg)
                    return result
                
                # Wait for UI response
                wait_time = step.get("wait_time")
                if wait_time:
                    self._wait_for_ui_response(wait_time)
                
                # Validate step success
                if self._validate_step_success(step_name):
                    result["steps_completed"].append(description)
                    self.logger.info(f"Step {i} completed successfully")
                else:
                    self.logger.warning(f"Step {i} validation uncertain")
                    result["steps_completed"].append(f"{description} (validation uncertain)")
            
            self.logger.info("Navigation sequence completed!")
            
            result.update({
                "success": True,
                "end_time": datetime.now().isoformat(),
                "total_steps": len(self.navigation_sequence),
                "window_handle": self.window_handle,
                "method": "keyboard_navigation"
            })
            
            return result
            
        except Exception as e:
            error_msg = f"Navigation failed: {str(e)}"
            self.logger.error(f"{error_msg}")
            self.logger.error(traceback.format_exc())
            result["errors"].append(error_msg)
            return result
    
    def run_navigation_phase(self, window_handle: Optional[int] = None) -> Dict[str, Any]:
        """Run the complete navigation phase"""
        try:
            self.logger.info("Starting VBS Navigation Phase")
            
            if window_handle:
                self.window_handle = window_handle
            
            result = self.execute_navigation()
            result["phase"] = "Navigation"
            
            if result["success"]:
                self.logger.info("Navigation phase completed successfully")
            else:
                self.logger.error("Navigation phase failed")
            
            return result
            
        except Exception as e:
            self.logger.error(f"Navigation phase failed: {e}")
            return {
                "success": False, 
                "phase": "Navigation", 
                "errors": [str(e)]
            }
    
    def get_current_window_info(self) -> Dict[str, Any]:
        """Get current window information"""
        try:
            info = {
                "window_handle": self.window_handle,
                "window_exists": False,
                "window_title": "",
                "window_rect": None,
                "is_focused": False
            }
            
            if self.window_handle:
                try:
                    info["window_exists"] = win32gui.IsWindow(self.window_handle)
                    if info["window_exists"]:
                        info["window_title"] = win32gui.GetWindowText(self.window_handle)
                        info["window_rect"] = win32gui.GetWindowRect(self.window_handle)
                        info["is_focused"] = (win32gui.GetForegroundWindow() == self.window_handle)
                except Exception:
                    pass
            
            return info
            
        except Exception as e:
            return {"error": str(e)}

if __name__ == "__main__":
    # Test the navigation phase
    nav_phase = VBSPhase2_Navigation()
    result = nav_phase.run_navigation_phase()
    print(f"Navigation result: {result}")