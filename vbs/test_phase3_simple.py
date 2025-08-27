#!/usr/bin/env python3
"""
Simple Phase 3 Test - Focus on Import Button Click
Test the critical import button clicking without audio detection
"""

import time
import logging
import pyautogui
import win32gui
from pathlib import Path

# Configure PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

class SimplePhase3Test:
    def __init__(self):
        self.logger = self._setup_logging()
        self.vbs_window = None
        self.images_dir = Path("Images/phase3")
        
        # Find VBS window
        self._find_vbs_window()
        
    def _setup_logging(self):
        logger = logging.getLogger("SimplePhase3")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        return logger
    
    def _find_vbs_window(self):
        """Find VBS window"""
        def check_window(hwnd, windows):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                title = win32gui.GetWindowText(hwnd).lower()
                if 'absons' in title or 'moonflower' in title:
                    windows.append((hwnd, title))
            except:
                pass
            return True
        
        windows = []
        win32gui.EnumWindows(check_window, windows)
        
        if windows:
            self.vbs_window = windows[0][0]
            self.logger.info(f"Found VBS window: {windows[0][1]}")
            return True
        else:
            self.logger.error("No VBS window found")
            return False
    
    def _focus_vbs(self):
        """Focus VBS window"""
        if self.vbs_window:
            try:
                win32gui.SetForegroundWindow(self.vbs_window)
                time.sleep(0.5)
                return True
            except:
                self.logger.warning("Could not focus VBS window")
                return False
        return False
    
    def _click_image_simple(self, image_name):
        """Simple image click with verification"""
        self.logger.info(f"Looking for: {image_name}")
        
        image_path = self.images_dir / image_name
        if not image_path.exists():
            self.logger.error(f"Image not found: {image_name}")
            return False
        
        self._focus_vbs()
        
        try:
            location = pyautogui.locateOnScreen(str(image_path), confidence=0.9)
            if location:
                click_x, click_y = pyautogui.center(location)
                pyautogui.click(click_x, click_y)
                self.logger.info(f"✅ Clicked {image_name} at ({click_x}, {click_y})")
                time.sleep(1.0)
                return True
            else:
                self.logger.error(f"❌ Could not find {image_name}")
                return False
        except Exception as e:
            self.logger.error(f"❌ Error clicking {image_name}: {e}")
            return False
    
    def test_sequence(self):
        """Test the critical sequence"""
        self.logger.info("🧪 TESTING PHASE 3 CRITICAL SEQUENCE")
        
        # Step 1: Import EHC checkbox
        self.logger.info("STEP 1: Import EHC checkbox")
        if not self._click_image_simple("01_import_ehc_checkbox.png"):
            return False
        
        # Step 2: Three dots button
        self.logger.info("STEP 2: Three dots button")
        if not self._click_image_simple("02_three_dots_button.png"):
            return False
        
        # Press ENTER immediately
        self.logger.info("STEP 2.1: Press ENTER for popup")
        pyautogui.press('enter')
        time.sleep(1.0)
        
        # Step 3: Check if we need to navigate or if import button is visible
        self.logger.info("STEP 3: Looking for import button directly")
        if self._click_image_simple("13_import_button.png"):
            self.logger.info("🎉 SUCCESS: Found and clicked import button!")
            return True
        else:
            self.logger.info("Import button not visible - need navigation first")
            return False

def main():
    """Test Phase 3 sequence"""
    print("🧪 SIMPLE PHASE 3 TEST - IMPORT BUTTON FOCUS")
    print("=" * 50)
    
    test = SimplePhase3Test()
    success = test.test_sequence()
    
    if success:
        print("✅ Phase 3 critical sequence working!")
    else:
        print("❌ Phase 3 sequence needs investigation")
    
    print("=" * 50)

if __name__ == "__main__":
    main() 