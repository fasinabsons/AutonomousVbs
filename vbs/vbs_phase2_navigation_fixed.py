#!/usr/bin/env python3
"""
VBS Phase 2 Navigation - Correct Image-Based Implementation
Sequence: Arrow → Sales & Distribution → POS → WiFi User Registration → New Button → Credit Radio Button
Uses OpenCV + PyAutoGUI for accurate image-based clicking to avoid minimizing issues
"""

import time
import logging
import os
from pathlib import Path
from datetime import datetime
import pyautogui
import win32gui
import win32con
import win32api

# Configure PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.2

class VBSPhase2Clean:
    """Clean VBS Phase 2 Navigation - Image-based clicking for all steps"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        self.vbs_window = None
        self.images_dir = None
        
        # Initialize
        self._find_vbs_window()
        self._setup_images()
        
        # Timing configuration
        self.delays = {
            "after_click": 1.5,        # Wait after each click
            "menu_open": 2.5,          # Menu opening time
            "form_load": 3.0,          # Form loading time
            "button_response": 1.0,    # Button response time
            "window_focus": 0.5        # Window focus time
        }
        
        self.logger.info("VBS Phase 2 Clean - initialized with image-based clicking")
    
    def _setup_logging(self):
        """Simple logging setup"""
        logger = logging.getLogger("VBSPhase2Clean")
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # Console only
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - VBS2 - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        
        return logger
    
    def _find_vbs_window(self):
        """Find ONLY VBS window - ignore everything else"""
        self.logger.info("Looking for VBS window...")
        
        def check_window(hwnd, windows):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                    
                title = win32gui.GetWindowText(hwnd).lower()
                
                # ONLY look for VBS indicators
                if any(word in title for word in ['absons', 'arabian', 'moonflower']):
                    # EXCLUDE non-VBS windows
                    if not any(word in title for word in ['sql', 'outlook', 'browser', 'chrome', 'firefox']):
                        windows.append((hwnd, title))
                        
            except:
                pass
            return True
        
        windows = []
        win32gui.EnumWindows(check_window, windows)
        
        if windows:
            self.vbs_window = windows[0][0]
            self.logger.info(f"✅ Found VBS window: {windows[0][1]}")
            return True
        else:
            self.logger.error("❌ No VBS window found")
            return False
    
    def _setup_images(self):
        """Setup image directory"""
        # Find Images/phase2 directory
        possible_paths = [
            Path("Images/phase2"),
            Path("../Images/phase2"),
            Path("../../Images/phase2")
        ]
        
        for path in possible_paths:
            if path.exists():
                self.images_dir = path
                break
        
        if self.images_dir:
            self.logger.info(f"✅ Images directory: {self.images_dir}")
            
            # Check for required images
            required_images = [
                "01_arrow_button.png",
                "02_sales_distribution_menu.png", 
                "03_pos_menu.png",
                "04_wifi_user_registration.png",
                "05_new_button.png",
                "06_credit_radio_button.png"
            ]
            
            missing_images = []
            for img in required_images:
                if not (self.images_dir / img).exists():
                    missing_images.append(img)
            
            if missing_images:
                self.logger.warning(f"⚠️ Missing images: {missing_images}")
            else:
                self.logger.info("✅ All required images found")
        else:
            self.logger.error("❌ Images/phase2 directory not found")
    
    def _focus_vbs_only(self):
        """SUPERIOR VBS window focusing - brings VBS to TOP above ALL applications"""
        if not self.vbs_window:
            return False
            
        try:
            # Make sure it's still valid
            if not win32gui.IsWindow(self.vbs_window):
                self._find_vbs_window()
                if not self.vbs_window:
                    return False
            
            # 🚀 SUPERIOR FOCUS STRATEGY: Multi-layer window elevation
            self.logger.debug(f"🎯 SUPERIOR FOCUS (Phase 2): Bringing VBS window {self.vbs_window} to absolute top")
            
            # STEP 1: Restore window if minimized/hidden
            if win32gui.IsIconic(self.vbs_window):
                win32gui.ShowWindow(self.vbs_window, win32con.SW_RESTORE)
                time.sleep(0.3)
            
            # STEP 1.5: Ensure window is MAXIMIZED (full screen) - prevent shrinking
            win32gui.ShowWindow(self.vbs_window, win32con.SW_MAXIMIZE)  # 🔥 ALWAYS FULL SCREEN
            time.sleep(0.2)
            
            # STEP 2: Force window to TOPMOST layer (above ALL apps)
            win32gui.SetWindowPos(
                self.vbs_window, 
                win32con.HWND_TOPMOST,  # 🔥 TOPMOST = Above everything
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            )
            time.sleep(0.2)
            
            # STEP 3: Bring to front and activate
            win32gui.BringWindowToTop(self.vbs_window)
            win32gui.SetForegroundWindow(self.vbs_window)
            time.sleep(0.2)
            
            # STEP 4: Use ctypes for additional enforcement
            import ctypes
            user32 = ctypes.windll.user32
            user32.SetForegroundWindow(self.vbs_window)
            user32.BringWindowToTop(self.vbs_window)
            time.sleep(0.2)
            
            # STEP 5: Return to normal layer but keep focus
            win32gui.SetWindowPos(
                self.vbs_window, 
                win32con.HWND_TOP,  # Normal top layer
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            )
            
            time.sleep(self.delays["window_focus"])
            self.logger.info("✅ SUPERIOR FOCUS: VBS window brought to top")
            return True
            
        except Exception as e:
            self.logger.error(f"SUPERIOR FOCUS failed: {e}")
            # Fallback to basic focus
            try:
                win32gui.SetForegroundWindow(self.vbs_window)
                return True
            except:
                return False
    
    def _click_image(self, image_name, timeout=30, required=True):
        """Click on image - VBS window only with anti-minimize protection"""
        if not self.images_dir:
            self.logger.error("No images directory")
            return False
            
        image_path = self.images_dir / image_name
        if not image_path.exists():
            if required:
                self.logger.error(f"Image not found: {image_name}")
            return False
        
        # Focus VBS first - gentle approach
        if not self._focus_vbs_only():
            return False
        
        self.logger.info(f"Looking for: {image_name}")
        
        # Try to find image with increased attempts for reliability
        max_attempts = timeout // 0.5  # Convert timeout to attempts (0.5s per attempt)
        for attempt in range(int(max_attempts)):
            try:
                # Take screenshot and look for image
                location = pyautogui.locateOnScreen(str(image_path), confidence=0.8)
                if location:
                    center = pyautogui.center(location)
                    
                    # Ensure VBS is still focused before clicking
                    self._focus_vbs_only()
                    time.sleep(0.3)
                    
                    # Gentle click to avoid window state changes
                    pyautogui.click(center)
                    self.logger.info(f"✅ Clicked {image_name} at {center}")
                    return True
                    
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                if required:
                    self.logger.warning(f"Click attempt {attempt+1} failed: {e}")
            
            time.sleep(0.5)
            
        if required:
            self.logger.error(f"❌ Could not find {image_name} after {timeout} seconds")
        return False
    
    def _press_key(self, key):
        """Press key in VBS window"""
        if not self._focus_vbs_only():
            return False
        
        key_codes = {
            "ENTER": win32con.VK_RETURN,
            "TAB": win32con.VK_TAB,
            "LEFT": win32con.VK_LEFT,
            "RIGHT": win32con.VK_RIGHT,
            "DOWN": win32con.VK_DOWN
        }
        
        if key not in key_codes:
            self.logger.error(f"Unknown key: {key}")
            return False
        
        try:
            vk_code = key_codes[key]
            win32api.keybd_event(vk_code, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
            self.logger.info(f"Pressed {key}")
            return True
            
        except Exception as e:
            self.logger.error(f"Key press failed: {e}")
            return False
    
    def execute_phase2(self):
        """Execute complete Phase 2 with correct image-based sequence"""
        self.logger.info("🚀 Starting VBS Phase 2 - Correct Image-Based Navigation")
        
        # Check prerequisites
        if not self.vbs_window:
            return {"success": False, "error": "No VBS window found"}
        
        if not self.images_dir:
            return {"success": False, "error": "No images directory"}
        
        steps = []
        
        try:
            # STEP 1: Click Arrow button
            self.logger.info("Step 1: Arrow button")
            if not self._click_image("01_arrow_button.png"):
                return {"success": False, "error": "Failed to click arrow button", "step": 1}
            steps.append("arrow_clicked")
            time.sleep(self.delays["menu_open"])
            
            # STEP 2: Click Sales & Distribution
            self.logger.info("Step 2: Sales & Distribution")
            if not self._click_image("02_sales_distribution_menu.png"):
                return {"success": False, "error": "Failed to click Sales & Distribution", "step": 2}
            steps.append("sales_clicked")
            time.sleep(self.delays["menu_open"])
            
            # STEP 3: Click POS
            self.logger.info("Step 3: POS menu")
            if not self._click_image("03_pos_menu.png"):
                return {"success": False, "error": "Failed to click POS", "step": 3}
            steps.append("pos_clicked")
            time.sleep(self.delays["menu_open"])
            
            # STEP 3.1: Press 3 DOWN arrow keys after POS to ensure correct position
            self.logger.info("Step 3.1: 3 DOWN arrow keys after POS for reliable navigation")
            for i in range(3):
                self._press_key("DOWN")
                time.sleep(0.2)
            steps.append("pos_navigation_arrows")
            
            # STEP 4: Navigate to WiFi User Registration using keyboard (RELIABLE METHOD)
            self.logger.info("Step 4: WiFi User Registration - KEYBOARD NAVIGATION")
            self.logger.info("Using keyboard: Already moved 3 DOWN from POS, now press ENTER")
            
            # We're already positioned at WiFi User Registration after the 3 DOWN arrows
            # Just press ENTER to select it
            if not self._press_key("ENTER"):
                return {"success": False, "error": "Failed to press ENTER for WiFi User Registration", "step": 4}
            
            steps.append("wifi_registration_selected")
            time.sleep(self.delays["form_load"])  # Wait for form to load
            
            # STEP 4b: WiFi Registration form should now be open
            self.logger.info("Step 4b: WiFi Registration form opened")
            steps.append("wifi_form_opened")
            
            # STEP 5: Click New Button
            self.logger.info("Step 5: New Button")
            if not self._click_image("05_new_button.png"):
                return {"success": False, "error": "Failed to click New button", "step": 5}
            steps.append("new_button_clicked")
            time.sleep(self.delays["button_response"])
            
            # STEP 6: Click Credit Radio Button
            self.logger.info("Step 6: Credit Radio Button")
            if not self._click_image("06_credit_radio_button.png"):
                return {"success": False, "error": "Failed to click Credit radio button", "step": 6}
            steps.append("credit_radio_selected")
            time.sleep(self.delays["after_click"])
            
            self.logger.info("🎉 SUCCESS: Phase 2 completed with all image clicks!")
            
            return {
                "success": True,
                "steps_completed": steps,
                "total_steps": len(steps),
                "message": "Phase 2 navigation completed successfully with image-based clicking",
                "sequence": "Arrow → Sales & Distribution → POS → WiFi User Registration → New Button → Credit Radio Button"
            }
            
        except Exception as e:
            self.logger.error(f"Phase 2 failed at step {len(steps)+1}: {e}")
            return {
                "success": False, 
                "error": str(e), 
                "steps_completed": steps,
                "failed_at_step": len(steps) + 1
            }

def main():
    """Test the corrected Phase 2"""
    print("Testing VBS Phase 2 - Correct Image-Based Navigation")
    print("=" * 60)
    print("Sequence: Arrow → Sales & Distribution → POS → WiFi User Registration → New → Credit")
    print("=" * 60)
    
    phase2 = VBSPhase2Clean()
    result = phase2.execute_phase2()
    
    print(f"\nResult: {result}")
    
    if result["success"]:
        print("✅ Phase 2 completed successfully!")
        print(f"✅ Steps completed: {len(result['steps_completed'])}/7")
        print(f"✅ Sequence: {result['sequence']}")
        print("\nCompleted steps:")
        for i, step in enumerate(result['steps_completed'], 1):
            print(f"  {i}. {step}")
    else:
        print("❌ Phase 2 failed")
        print(f"❌ Error: {result['error']}")
        if 'failed_at_step' in result:
            print(f"❌ Failed at step: {result['failed_at_step']}")
        if 'steps_completed' in result:
            print(f"✅ Steps completed before failure: {result['steps_completed']}")

def main():
    """Test the corrected Phase 2"""
    print("Testing VBS Phase 2 - Correct Image-Based Navigation")
    print("=" * 60)
    print("Sequence: Arrow → Sales & Distribution → POS → WiFi User Registration → New → Credit")
    print("=" * 60)
    
    phase2 = VBSPhase2Clean()
    result = phase2.execute_phase2()
    
    print(f"\nResult: {result}")
    
    if result["success"]:
        print("✅ Phase 2 completed successfully!")
        print(f"✅ Steps completed: {len(result['steps_completed'])}/7")
        print(f"✅ Sequence: {result['sequence']}")
        print("\nCompleted steps:")
        for i, step in enumerate(result['steps_completed'], 1):
            print(f"  {i}. {step}")
    else:
        print("❌ Phase 2 failed")
        print(f"❌ Error: {result['error']}")
        if 'failed_at_step' in result:
            print(f"❌ Failed at step: {result['failed_at_step']}")
        if 'steps_completed' in result:
            print(f"✅ Steps completed before failure: {result['steps_completed']}")
    
    return result

if __name__ == "__main__":
    import sys
    result = main()
    
    # Exit with proper code for BAT file compatibility
    if result and result.get("success", False):
        print("\n🎉 Phase 2 completed successfully - exiting with success code")
        sys.exit(0)  # Success
    else:
        print("\n❌ Phase 2 failed - exiting with error code")
        sys.exit(1)  # Failure