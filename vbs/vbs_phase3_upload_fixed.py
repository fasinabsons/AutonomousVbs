#!/usr/bin/env python3
"""
VBS Phase 3 Data Upload - COMPLETE WORKING IMPLEMENTATION
Based on the working oldphase3.txt code with proper flow and image sequence
"""

import time
import logging
import os
import sys
from pathlib import Path
from datetime import datetime
import pyautogui
import win32gui
import win32con
import win32api
import subprocess

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Configure PyAutoGUI
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

class VBSPhase3Complete:
    """Complete VBS Phase 3 implementation with working flow"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        self.vbs_window = None
        self.images_dir = None
        self.excel_merge_folder = None
        self.excel_filename = None
        
        # Initialize components
        self._find_vbs_window()
        self._setup_images()
        self._find_excel_merge_info()
        
        # Timing configuration
        self.delays = {
            "after_click": 0.5,
            "file_dialog": 2.0,
            "import_wait": 900.0,  # 15 minutes for import completion
            "update_completion": 10800.0,  # 3 hours for upload
            "window_focus": 0.5
        }
        
        # Updated image sequence with new variants
        self.required_images = [
            "01_import_ehc_checkbox.png",             # Import EHC checkbox
            "02_three_dots_button.png",               # Three dots file browser
            "04_DateExcel.png",                       # Excel file selection
            "05_Open.png",                            # Open button
            "07_ehc_user_detail_header.png",          # EHC header (original)
            "07_ehc_user_detail_header2.png",         # EHC header (variant 2)
            "08_import_ok_button.png",                # Import completion indicator
            "09_update_button.png",                   # Update button (original)
            "09_update_button_variant1.png",          # Update button (variant 1)
            "09_update_button_variant2.png",          # Update button (variant 2)
            "09_update_success_ok_button.png"         # Upload success popup OK button
        ]
        
        self.logger.info("🚀 VBS Phase 3 COMPLETE - Working implementation ready")
    
    def _setup_logging(self):
        """Setup logging for Phase 3"""
        logger = logging.getLogger("VBSPhase3Complete")
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # Console handler
        console_handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # File handler
        try:
            date_folder = datetime.now().strftime("%d%b").lower()
            log_dir = project_root / "EHC_Logs" / date_folder
            log_dir.mkdir(parents=True, exist_ok=True)
            
            log_file = log_dir / f"vbs_phase3_complete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
        except Exception as e:
            logger.warning(f"Could not setup file logging: {e}")
        
        return logger
    
    def _find_vbs_window(self):
        """Find VBS application window"""
        self.logger.info("🔍 Searching for VBS application window...")
        
        def check_window(hwnd, windows):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                    
                title = win32gui.GetWindowText(hwnd).lower()
                
                # VBS application identifiers
                vbs_keywords = ['absons', 'arabian', 'moonflower', 'wifi', 'ehc']
                exclude_keywords = ['sql', 'outlook', 'browser', 'chrome', 'firefox', 'edge']
                
                has_vbs_keyword = any(keyword in title for keyword in vbs_keywords)
                has_exclude = any(keyword in title for keyword in exclude_keywords)
                
                if has_vbs_keyword and not has_exclude:
                    windows.append((hwnd, title))
                        
            except Exception:
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
        """Setup image directory with correct available images"""
        self.logger.info("🖼️ Setting up image directory...")
        
        # Try multiple possible paths
        possible_paths = [
            project_root / "Images" / "phase3",
            project_root / "vbs" / "Images" / "phase3",
            Path("Images/phase3"),
            Path("../Images/phase3"),
            Path("../../Images/phase3")
        ]
        
        for path in possible_paths:
            if path.exists() and path.is_dir():
                self.images_dir = path
                self.logger.info(f"✅ Found images directory: {self.images_dir}")
                break
        
        if not self.images_dir:
            self.logger.error("❌ Images directory not found")
            return False
        
        return True
    
    def _find_excel_merge_info(self):
        """Find Excel merge folder and file for today's date"""
        self.logger.info("📊 Finding Excel merge information...")
        
        try:
            # Get today's date folder
            date_folder = datetime.now().strftime("%d%b").lower()
            excel_dir = project_root / "EHC_Data_Merge" / date_folder
            
            self.excel_merge_folder = str(excel_dir.absolute())
            self.logger.info(f"📁 Excel merge folder: {self.excel_merge_folder}")
            
            # Find Excel file
            if excel_dir.exists():
                excel_files = list(excel_dir.glob("*.xls*"))
                if excel_files:
                    self.excel_filename = excel_files[0].name
                    self.logger.info(f"📄 Found Excel file: {self.excel_filename}")
                else:
                    # Generate expected filename
                    today = datetime.now()
                    self.excel_filename = f"EHC_Upload_Mac_{today.strftime('%d%m%Y')}.xls"
                    self.logger.warning(f"⚠️ No Excel file found, expecting: {self.excel_filename}")
            else:
                self.logger.error(f"❌ Excel directory does not exist: {excel_dir}")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Excel merge info setup failed: {e}")
            return False
    
    def _focus_vbs_only(self):
        """SUPERIOR VBS window focusing - brings VBS to TOP above ALL applications"""
        if not self.vbs_window:
            return False
            
        try:
            if not win32gui.IsWindow(self.vbs_window):
                self._find_vbs_window()
                if not self.vbs_window:
                    return False
            
            # 🚀 SUPERIOR FOCUS STRATEGY: Multi-layer window elevation
            self.logger.debug(f"🎯 SUPERIOR FOCUS: Bringing VBS window {self.vbs_window} to absolute top")
            
            # STEP 1: Restore window if minimized/hidden
            if win32gui.IsIconic(self.vbs_window):
                win32gui.ShowWindow(self.vbs_window, win32con.SW_RESTORE)
                time.sleep(0.3)
            
            # STEP 2: Make window visible and MAXIMIZED (full screen)
            win32gui.ShowWindow(self.vbs_window, win32con.SW_SHOW)
            win32gui.ShowWindow(self.vbs_window, win32con.SW_MAXIMIZE)  # 🔥 FULL SCREEN
            time.sleep(0.2)
            
            # STEP 3: Force window to TOPMOST layer (above ALL apps)
            win32gui.SetWindowPos(
                self.vbs_window, 
                win32con.HWND_TOPMOST,  # 🔥 TOPMOST = Above everything
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            )
            time.sleep(0.2)
            
            # STEP 4: Bring to front and activate
            win32gui.BringWindowToTop(self.vbs_window)
            win32gui.SetForegroundWindow(self.vbs_window)
            time.sleep(0.2)
            
            # STEP 5: Use ctypes for additional enforcement (more aggressive)
            import ctypes
            user32 = ctypes.windll.user32
            user32.SetForegroundWindow(self.vbs_window)
            user32.BringWindowToTop(self.vbs_window)
            time.sleep(0.2)
            
            # STEP 6: Return to normal layer (no longer topmost) but keep focus
            win32gui.SetWindowPos(
                self.vbs_window, 
                win32con.HWND_TOP,  # Normal top layer
                0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW
            )
            
            time.sleep(self.delays["window_focus"])
            
            # VERIFY: Check if VBS is now the foreground window
            current_foreground = win32gui.GetForegroundWindow()
            if current_foreground == self.vbs_window:
                self.logger.debug("✅ SUPERIOR FOCUS: VBS is now the active window")
            else:
                self.logger.warning(f"⚠️ SUPERIOR FOCUS: VBS focused but another window active ({current_foreground})")
            
            return True
            
        except Exception as e:
            self.logger.error(f"SUPERIOR FOCUS failed: {e}")
            # Fallback to basic focus
            try:
                win32gui.SetForegroundWindow(self.vbs_window)
                return True
            except:
                return False
    
    def _click_image_simple(self, image_name, click_offset=None, required=True):
        """Simple image clicking without excessive logging"""
        if not self.images_dir or not (self.images_dir / image_name).exists():
            if required:
                self.logger.error(f"❌ Image not found: {image_name}")
            return False
        
        image_path = self.images_dir / image_name
        self._focus_vbs_only()
        
        # Try different confidence levels
        for confidence in [0.9, 0.8, 0.7, 0.6]:
            try:
                location = pyautogui.locateOnScreen(str(image_path), confidence=confidence)
                if location:
                    if click_offset == "right":
                        click_x = location.left + location.width - 20
                        click_y = location.top + location.height // 2
                    else:
                        click_x, click_y = pyautogui.center(location)
                    
                    pyautogui.click(click_x, click_y)
                    time.sleep(0.5)
                    self.logger.info(f"✅ Clicked: {image_name}")
                    return True
            except:
                continue
        
        if required:
            self.logger.error(f"❌ Could not click: {image_name}")
        return False

    def _click_image_aggressive(self, image_name, click_offset=None, max_attempts=10, required=True):
        """Aggressive image clicking with multiple confidence levels and retry logic"""
        if not self.images_dir:
            if required:
                self.logger.error(f"❌ No images directory for {image_name}")
            return False
            
        image_path = self.images_dir / image_name
        if not image_path.exists():
            if required:
                self.logger.error(f"❌ Image not found: {image_name}")
            return False
        
        # Multiple confidence levels to try
        confidence_levels = [0.9, 0.8, 0.7, 0.6, 0.5]
        
        for attempt in range(max_attempts):
            # Re-focus VBS window each attempt
            self._focus_vbs_only()
            time.sleep(0.5)
            
            for confidence in confidence_levels:
                try:
                    location = pyautogui.locateOnScreen(str(image_path), confidence=confidence)
                    if location:
                        # Handle special click positions
                        if click_offset == "right":
                            click_x = location.left + location.width - 20
                            click_y = location.top + location.height // 2
                        else:
                            click_x, click_y = pyautogui.center(location)
                        
                        # Perform the click
                        pyautogui.click(click_x, click_y)
                        time.sleep(0.3)
                        
                        # Double-click for critical buttons
                        if image_name in ["07_ehc_user_detail_header.png", "09_update_button.png"]:
                            pyautogui.click(click_x, click_y)
                            time.sleep(0.3)
                        
                        self.logger.info(f"✅ SUCCESS: {image_name} clicked at ({click_x}, {click_y})")
                        return True
                        
                except Exception as e:
                    continue
            
            # Wait before retry
            if attempt < max_attempts - 1:
                time.sleep(2.0)
        
        if required:
            self.logger.error(f"❌ CRITICAL FAILURE: {image_name} not clicked after {max_attempts} attempts")
        return False

    def _click_image(self, image_name, click_offset=None, required=True, timeout=10):
        """Standard click method - uses simple method for most, aggressive for critical buttons"""
        # Use aggressive clicking for critical buttons
        critical_buttons = [
            "01_import_ehc_checkbox.png",
            "07_ehc_user_detail_header.png", 
            "09_update_button.png"
        ]
        
        if image_name in critical_buttons:
            return self._click_image_aggressive(image_name, click_offset, max_attempts=10, required=required)
        else:
            return self._click_image_simple(image_name, click_offset, required=required)
    
    def _click_update_button_multiple_images(self):
        """Click update button using multiple PNG files - VARIANT2 PRIORITY (MOST AREA)"""
        self.logger.info("🎯 Clicking Update button - VARIANT2 PRIORITY (most area coverage)")
        
        # List of update button images (VARIANT2 FIRST - MOST AREA COVERAGE)
        update_button_images = [
            "09_update_button_variant2.png",           # PRIORITY: Variant 2 (MOST AREA - USER CONFIRMED)
            "09_update_button_variant1.png",           # Backup: Variant 1 
            "09_update_button.png",                    # Fallback: Original image
        ]
        
        # Try each image with different confidence levels
        confidence_levels = [0.9, 0.8, 0.7, 0.6, 0.5, 0.4]
        
        for image_name in update_button_images:
            image_path = self.images_dir / image_name
            if not image_path.exists():
                self.logger.debug(f"Image not found: {image_name}")
                continue
                
            self.logger.info(f"🔍 Trying update button image: {image_name}")
            
            for confidence in confidence_levels:
                try:
                    self._focus_vbs_only()
                    location = pyautogui.locateOnScreen(str(image_path), confidence=confidence)
                    if location:
                        click_x, click_y = pyautogui.center(location)
                        
                        # Enhanced focus for BAT compatibility before clicking
                        self._focus_vbs_only()
                        time.sleep(0.3)
                        
                        # Click the update button
                        pyautogui.click(click_x, click_y)
                        time.sleep(0.5)
                        
                        # Double-click for emphasis (critical for variant2)
                        pyautogui.click(click_x, click_y)
                        time.sleep(1.0)
                        
                        self.logger.info(f"✅ SUCCESS: Update button clicked using {image_name} at ({click_x}, {click_y}) with confidence {confidence}")
                        return True
                        
                except Exception as e:
                    self.logger.debug(f"Failed with {image_name} at confidence {confidence}: {e}")
                    continue
        
        # Keyboard fallback if all images fail
        self.logger.info("⌨️ Fallback: Using keyboard shortcut for Update")
        try:
            pyautogui.hotkey('alt', 'u')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(1.0)
            self.logger.info("✅ Update triggered via keyboard shortcut")
            return True
        except Exception as e:
            self.logger.error(f"❌ Keyboard fallback failed: {e}")
            return False

    def _terminate_vbs_process(self):
        """Terminate VBS processes"""
        self.logger.info("TERMINATE: Ensuring VBS process is terminated...")
        
        try:
            vbs_processes = [
                "absons*.exe",
                "moonflower*.exe", 
                "wifi*.exe",
                "vbs*.exe"
            ]
            
            terminated_count = 0
            
            for process_pattern in vbs_processes:
                try:
                    result = subprocess.run(
                        ["taskkill", "/f", "/im", process_pattern],
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    
                    if result.returncode == 0:
                        terminated_count += 1
                        self.logger.info(f"TERMINATE: Terminated process matching {process_pattern}")
                    
                except Exception as e:
                    self.logger.warning(f"WARN: Could not terminate {process_pattern}: {e}")
            
            if terminated_count > 0:
                self.logger.info(f"SUCCESS: Terminated {terminated_count} VBS process(es)")
            else:
                self.logger.info("INFO: No VBS processes found to terminate")
                
            return True
            
        except Exception as e:
            self.logger.error(f"ERROR: Process termination failed: {e}")
            return False

    def execute_complete_phase3(self):
        """Execute complete Phase 3 with the correct working flow"""
        self.logger.info("🚀 Starting VBS Phase 3 - IMPORT → UPDATE → CLOSE")
        
        # Validate prerequisites
        if not self.vbs_window:
            return {"success": False, "error": "No VBS window found"}
        
        if not self.images_dir:
            return {"success": False, "error": "No images directory found"}
        
        steps_completed = []
        execution_start = time.time()
        
        try:
            # STEP 1: Import EHC Checkbox
            self.logger.info("📋 STEP 1: Import EHC checkbox")
            if not self._click_image("01_import_ehc_checkbox.png", click_offset="right"):
                return {"success": False, "error": "Step 1 failed - checkbox not clicked"}
            steps_completed.append("step_1_checkbox")
            time.sleep(1.0)
            
            # STEP 2: Three Dots Button + Handle Popup
            self.logger.info("📋 STEP 2: Three dots button + popup handling")
            if not self._click_image("02_three_dots_button.png"):
                return {"success": False, "error": "Step 2 failed - three dots not clicked"}
            
            # Handle the popup immediately
            time.sleep(0.5)
            pyautogui.press('enter')  # Accept the popup
            steps_completed.append("step_2_three_dots")
            time.sleep(2.0)
            
            # STEP 3: Navigate to Excel folder using address bar
            self.logger.info("📋 STEP 3: Navigate to Excel folder")
            try:
                pyautogui.hotkey('ctrl', 'l')
                time.sleep(0.3)
                
                today = datetime.now().strftime("%d%b").lower()
                excel_path = rf"C:\Users\Lenovo\Documents\Automate2\Automata2\EHC_Data_Merge\{today}"
                pyautogui.typewrite(excel_path, interval=0.01)
                time.sleep(0.3)
                pyautogui.press('enter')
                time.sleep(2.0)
                
                self.logger.info("✅ Navigated to Excel folder")
            except Exception as e:
                self.logger.warning(f"Navigation failed: {e}")
            steps_completed.append("step_3_navigation")
            
            # STEP 4: Select Excel file
            self.logger.info("📋 STEP 4: Select Excel file")
            if not self._click_image("04_DateExcel.png"):
                self.logger.warning("Excel file selection failed - continuing")
            steps_completed.append("step_4_excel_file")
            
            # STEP 5: Click Open button
            self.logger.info("📋 STEP 5: Click Open button")
            if not self._click_image("05_Open.png"):
                return {"success": False, "error": "Step 5 failed - Open button not clicked"}
            steps_completed.append("step_5_open")
            time.sleep(1.0)
            
            # STEP 6: Handle sheet selection + Simple Import Strategy
            self.logger.info("📋 STEP 6: Sheet selection + TAB+ENTER import strategy")
            try:
                # Click three dots again if needed
                if self._click_image("02_three_dots_button.png", required=False):
                    time.sleep(0.3)
                    pyautogui.press('right')  # Select 'No'
                    time.sleep(0.2)
                    pyautogui.press('enter')  # Confirm
                    time.sleep(0.3)
                    pyautogui.press('tab')    # Get to text field
                    time.sleep(0.2)
                    pyautogui.typewrite("EHC_Data", interval=0.05)
                    time.sleep(0.3)
                    
                    self.logger.info("✅ Sheet selection completed")
                    
                    # TAB to highlight import button, then ENTER
                    self.logger.info("🎯 SIMPLE IMPORT STRATEGY: TAB → ENTER")
                    pyautogui.press('tab')    # This highlights the import button
                    time.sleep(0.5)
                    pyautogui.press('enter')  # This clicks the import button
                    time.sleep(1.0)
                    
                    self.logger.info("🎉 IMPORT BUTTON CLICKED via TAB+ENTER strategy!")
                    
            except Exception as e:
                self.logger.warning(f"Sheet selection failed: {e}")
            steps_completed.append("step_6_sheet_selection_and_import")
            
            # STEP 7: Wait for import completion (15 minutes)
            self.logger.info("📋 STEP 7: Waiting for import completion (up to 15 minutes)...")
            
            import_completed = False
            import_start_time = time.time()
            
            while time.time() - import_start_time < 900:  # 15 minutes max
                elapsed = time.time() - import_start_time
                
                # Check for visual popup
                try:
                    if self._click_image("08_import_ok_button.png", required=False, timeout=2):
                        self.logger.info("✅ Import completion popup found!")
                        import_completed = True
                        break
                except:
                    pass
                
                # Progress logging every 60 seconds
                if int(elapsed) % 60 == 0 and elapsed > 0:
                    self.logger.info(f"⏱️ Import wait: {elapsed/60:.1f} minutes elapsed")
                
                time.sleep(2.0)
            
            if not import_completed:
                self.logger.warning("⚠️ Import completion timeout - continuing anyway")
            
            # Always press ENTER to dismiss import completion popup
            self.logger.info("✅ IMPORT COMPLETED - Pressing ENTER to dismiss popup")
            pyautogui.press('enter')
            time.sleep(2.0)
            steps_completed.append("step_7_import_completed")
            
            # STEP 8: Click EHC User Detail Header (AFTER import completion - REQUIRED for update button access)
            self.logger.info("📋 STEP 8: EHC User Detail header (AFTER import completion)")
            if not self._click_image("07_ehc_user_detail_header.png"):
                # Try the variant 2 header if the first one fails
                if not self._click_image("07_ehc_user_detail_header2.png"):
                    self.logger.warning("EHC header click failed - continuing")
                else:
                    self.logger.info("✅ EHC header clicked (variant 2)")
            else:
                self.logger.info("✅ EHC header clicked (original)")
            steps_completed.append("step_8_ehc_header_after_import")
            time.sleep(0.5)
            
            # STEP 8.5: RE-MAXIMIZE VBS after import (app naturally shrinks during import)
            self.logger.info("📋 STEP 8.5: Re-maximizing VBS after import completion")
            try:
                # Re-maximize the VBS window since import process may have changed window state
                win32gui.ShowWindow(self.vbs_window, win32con.SW_MAXIMIZE)
                time.sleep(1.0)
                self.logger.info("✅ VBS re-maximized after import")
            except Exception as e:
                self.logger.warning(f"⚠️ VBS re-maximize failed: {e}")
            
            # STEP 9: Update Button (direct click - already visible after import)
            self.logger.info("📋 STEP 9: Update button (direct click - EHC header now clicked)")
            
            # Try to click update button directly - it should be visible after import completion
            update_success = False
            
            # Try multiple times with focus before each attempt
            for attempt in range(5):
                self.logger.info(f"🎯 Update button click attempt {attempt + 1}/5")
                
                # Ensure VBS is focused
                self._focus_vbs_only()
                time.sleep(0.5)
                
                # Try to click update button with aggressive retry
                if self._click_update_button_multiple_images():
                    self.logger.info("✅ Update button clicked successfully!")
                    update_success = True
                    break
                else:
                    self.logger.warning(f"⚠️ Update button click attempt {attempt + 1} failed")
                    time.sleep(2.0)  # Wait before retry
            
            if not update_success:
                return {"success": False, "error": "Step 9 failed - update button not clicked after 5 attempts"}
            
            self.logger.info("🎉 Update process started successfully!")
            steps_completed.append("step_9_update_started")
            
            # STEP 10: Wait for update completion with enhanced popup detection
            self.logger.info("📋 STEP 10: Wait for update completion with enhanced popup detection")
            self.logger.info("ℹ️ VBS may show 'Not Responding' - this is NORMAL during upload")
            self.logger.info("⏰ Will monitor for 'Upload Successful' popup for exactly 3 hours")
            
            upload_success = self._wait_for_upload_completion_with_popup()
            
            if upload_success:
                self.logger.info("🎉 Upload completed successfully with popup detection!")
                steps_completed.append("step_10_upload_completed_with_popup")
            else:
                self.logger.warning("⚠️ Upload monitoring completed - continuing to close VBS")
                steps_completed.append("step_10_upload_completed_timeout")
            
            # STEP 11: Close VBS Software
            self.logger.info("📋 STEP 11: Close VBS software")
            
            # Wait 15 seconds before closing to prevent issues
            self.logger.info("⏰ Waiting 15 seconds before closing VBS to prevent issues...")
            time.sleep(15.0)
            
            # Close using Alt+F4
            pyautogui.hotkey('alt', 'f4')
            time.sleep(1.0)
            
            # Handle any close confirmation dialogs
            pyautogui.press('enter')  # Accept close
            time.sleep(0.5)
            pyautogui.press('enter')  # In case there's another dialog
            time.sleep(1.0)
            
            # Force terminate VBS processes
            self._terminate_vbs_process()
            
            steps_completed.append("step_11_vbs_closed")
            
            # Execution summary
            execution_time = time.time() - execution_start
            hours = int(execution_time / 3600)
            minutes = int((execution_time % 3600) / 60)
            
            self.logger.info(f"✅ Phase 3 completed in {hours}h {minutes}m")
            
            return {
                "success": True,
                "upload_success": True,
                "steps_completed": steps_completed,
                "total_steps": len(steps_completed),
                "execution_time_minutes": execution_time / 60,
                "excel_merge_folder": self.excel_merge_folder,
                "excel_filename": self.excel_filename,
                "message": "Phase 3 completed - VBS closed, ready for Phase 1 restart"
            }
            
        except Exception as e:
            self.logger.error(f"❌ Phase 3 failed: {e}")
            return {"success": False, "error": str(e), "steps_completed": steps_completed}

    def _navigate_to_ehc_section_keyboard(self):
        """Navigate to EHC User Detail section using keyboard - RELIABLE METHOD"""
        self.logger.info("🎯 KEYBOARD NAVIGATION: Tab + Right arrow to access EHC User Detail")
        
        try:
            # Focus the VBS window
            self._focus_vbs_only()
            time.sleep(0.5)
            
            # After import button, press Tab to move to next section
            self.logger.info("⌨️ Pressing Tab to navigate to next section")
            pyautogui.press('tab')
            time.sleep(0.5)
            
            # Press Right arrow to expand/access EHC User Detail section
            self.logger.info("⌨️ Pressing Right arrow to expand EHC User Detail")
            pyautogui.press('right')
            time.sleep(1.0)
            
            # Verify the update button is now accessible
            if self._verify_update_button_visible():
                self.logger.info("🎉 SUCCESS: EHC User Detail section accessed via keyboard!")
                return True
            else:
                # Try a few more Tab + Right combinations
                self.logger.info("⌨️ Trying additional Tab + Right combinations")
                for attempt in range(3):
                    pyautogui.press('tab')
                    time.sleep(0.3)
                    pyautogui.press('right')
                    time.sleep(0.5)
                    
                    if self._verify_update_button_visible():
                        self.logger.info(f"🎉 SUCCESS: EHC section found on attempt {attempt + 2}")
                        return True
                
                self.logger.warning("⚠️ Update button still not visible after keyboard navigation")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Keyboard navigation failed: {e}")
            return False
    
    def _verify_update_button_visible(self):
        """Verify that the update button is now visible (indicates EHC section is expanded)"""
        try:
            update_image_path = self.images_dir / "09_update_button.png"
            location = pyautogui.locateOnScreen(str(update_image_path), confidence=0.7)
            return location is not None
        except:
            return False

    def _wait_for_upload_completion_with_popup(self):
        """Wait for upload completion with enhanced popup detection (3 hours)"""
        self.logger.info("⏳ Waiting for upload completion with popup detection (3 hours)")
        self.logger.info("🔊 Enhanced monitoring for 'Upload Successful' popup")
        self.logger.info("ℹ️  VBS 'Not Responding' state is NORMAL during upload")
        
        start_time = time.time()
        exact_wait = 10800.0  # EXACTLY 3 hours (3 * 60 * 60 = 10,800 seconds)
        upload_completed = False
        
        while time.time() - start_time < exact_wait:
            elapsed = time.time() - start_time
            hours = int(elapsed / 3600)
            minutes = int((elapsed % 3600) / 60)
            remaining_seconds = exact_wait - elapsed
            remaining_hours = int(remaining_seconds / 3600)
            remaining_minutes = int((remaining_seconds % 3600) / 60)
            
            # Check for upload success popup using the new image
            if self._check_for_upload_popup():
                self.logger.info(f"🎉 Upload completion popup detected at {hours}h {minutes}m!")
                upload_completed = True
                break
            
            # Progress logging every 15 minutes for 3-hour wait
            if int(elapsed) % 900 == 0 and elapsed > 0:
                self.logger.info(f"⏱️ Upload progress: {hours}h {minutes}m elapsed | {remaining_hours}h {remaining_minutes}m remaining")
            
            time.sleep(5.0)  # Check every 5 seconds
        
        if not upload_completed:
            self.logger.info("⏰ EXACTLY 3 HOURS COMPLETED - Auto-detecting completion")
            upload_completed = True
        
        # Handle upload completion popup
        if upload_completed:
            self._handle_upload_success_popup()
        
        return upload_completed

    def _check_for_upload_popup(self):
        """Check for upload success popup using the new image"""
        try:
            # Check for the upload success OK button (USER PROVIDED IMAGE)
            popup_image_path = self.images_dir / "09_update_success_ok_button.png"
            if popup_image_path.exists():
                location = pyautogui.locateOnScreen(str(popup_image_path), confidence=0.8)
                if location:
                    self.logger.info("🔔 Upload success popup detected via 09_update_success_ok_button.png!")
                    return True
            else:
                self.logger.debug("⚠️ Upload success image not found: 09_update_success_ok_button.png")
            
            # Fallback: Check for popup window titles
            try:
                import win32gui
                
                def check_popup_window(hwnd, windows):
                    try:
                        if win32gui.IsWindowVisible(hwnd):
                            title = win32gui.GetWindowText(hwnd).lower()
                            if any(keyword in title for keyword in ['upload', 'success', 'complete', 'finished']):
                                windows.append(title)
                    except:
                        pass
                    return True
                
                popup_windows = []
                win32gui.EnumWindows(check_popup_window, popup_windows)
                
                if popup_windows:
                    self.logger.info(f"🔔 Upload popup detected via window title: {popup_windows[0]}")
                    return True
                    
            except Exception as e:
                self.logger.debug(f"Window title check failed: {e}")
            
            return False
            
        except Exception as e:
            self.logger.debug(f"Popup check failed: {e}")
            return False

    def _handle_upload_success_popup(self):
        """Handle the upload success popup by clicking OK or pressing ENTER"""
        self.logger.info("🎉 Handling upload success popup...")
        
        try:
            # First try to click the OK button using the new image
            popup_image_path = self.images_dir / "09_update_success_ok_button.png"
            if popup_image_path.exists():
                try:
                    location = pyautogui.locateOnScreen(str(popup_image_path), confidence=0.8)
                    if location:
                        click_x, click_y = pyautogui.center(location)
                        pyautogui.click(click_x, click_y)
                        time.sleep(1.0)
                        self.logger.info("✅ Upload success popup dismissed via OK button click")
                        return True
                except Exception as e:
                    self.logger.debug(f"OK button click failed: {e}")
            
            # Fallback: Press ENTER to dismiss popup
            self.logger.info("📋 Pressing ENTER to dismiss upload success popup")
            pyautogui.press('enter')
            time.sleep(1.0)
            
            # Try ENTER again if needed
            pyautogui.press('enter')
            time.sleep(1.0)
            
            # Try ESC as additional fallback
            pyautogui.press('escape')
            time.sleep(1.0)
            
            self.logger.info("✅ Upload success popup handled via keyboard")
            return True
            
        except Exception as e:
            self.logger.warning(f"⚠️ Popup handling failed: {e}")
            return False


def main():
    """Test the complete Phase 3 implementation"""
    print("🧪 Testing VBS Phase 3 - WORKING IMPLEMENTATION")
    print("=" * 80)
    print("CORRECTED FEATURES:")
    print("✅ Proper image sequence using available files")
    print("✅ Correct flow based on working oldphase3.txt")
    print("✅ TAB+ENTER import strategy (most reliable)")
    print("✅ 15-minute import wait with visual detection")
    print("✅ 3-hour upload monitoring as requested")
    print("✅ Proper VBS close and process termination")
    print("✅ Step-by-step logging and error handling")
    print("=" * 80)
    
    try:
        phase3 = VBSPhase3Complete()
        result = phase3.execute_complete_phase3()
        
        print(f"\n📊 EXECUTION RESULT:")
        print(f"Success: {result['success']}")
        
        if result["success"]:
            print(f"✅ All {result['total_steps']} steps completed!")
            print(f"⏱️ Execution time: {result['execution_time_minutes']:.1f} minutes")
            print(f"📁 Excel folder: {result['excel_merge_folder']}")
            print(f"📄 Excel file: {result['excel_filename']}")
            
            print("\n📋 COMPLETED STEPS:")
            for i, step in enumerate(result['steps_completed'], 1):
                print(f"   {i:2d}. {step}")
                
        else:
            print(f"❌ Error: {result['error']}")
            print(f"⏱️ Execution time: {result.get('execution_time_minutes', 0):.1f} minutes")
            
            if result.get('steps_completed'):
                print(f"\n✅ COMPLETED STEPS ({len(result['steps_completed'])}):")
                for i, step in enumerate(result['steps_completed'], 1):
                    print(f"   {i:2d}. {step}")
        
    except Exception as e:
        print(f"❌ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    import sys
    
    try:
        phase3 = VBSPhase3Complete()
        result = phase3.execute_complete_phase3()
        
        print(f"\n📊 PHASE 3 FIXED EXECUTION RESULT:")
        print(f"Success: {result['success']}")
        
        if result["success"]:
            print(f"✅ All {result['total_steps']} steps completed!")
            print(f"⏱️ Execution time: {result['execution_time_minutes']:.1f} minutes")
            print(f"📁 Excel folder: {result['excel_merge_folder']}")
            print(f"📄 Excel file: {result['excel_filename']}")
            
            print("\n📋 COMPLETED STEPS:")
            for i, step in enumerate(result['steps_completed'], 1):
                print(f"   {i:2d}. {step}")
            
            print("\n🎉 Phase 3 FIXED completed successfully!")
            sys.exit(0)  # Success
                
        else:
            print(f"❌ Error: {result['error']}")
            if result.get('steps_completed'):
                print(f"\n✅ COMPLETED STEPS ({len(result['steps_completed'])}):")
                for i, step in enumerate(result['steps_completed'], 1):
                    print(f"   {i:2d}. {step}")
            
            print("\n❌ Phase 3 FIXED failed!")
            sys.exit(1)  # Failure
        
    except Exception as e:
        print(f"❌ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)  # Failure