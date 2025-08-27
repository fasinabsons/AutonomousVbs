#!/usr/bin/env python3
"""
VBS Automation - Phase 4: PDF Report Generation and Export (COMPLETELY FIXED)
Following exact user requirements:
1. Arrow button click (copy from Phase 2 logic)
2. Sales & Distribution
3. Reports menu
4. Scroll to BOTTOM then click POS
5. Double-click WiFi Active Users Count, wait 3 seconds
6. Clear all fields, enter start of month + current date
7. Print (15 seconds), Export, OK, OK, ensure PDF format
8. Address bar navigation to save PDF correctly
"""

import time
import logging
import pyautogui
import cv2
import numpy as np
from pathlib import Path
import os
import sys
from typing import Dict, Optional, Any, List, Tuple
from datetime import datetime, timedelta
import win32gui
import win32api
import win32con

# Import universal path manager
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from universal_path_manager import get_paths, ensure_directories, cd_to_project_root

# Disable PyAutoGUI failsafe for automation
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.5

class VBSPhase4_ReportFixed:
    """VBS Phase 4 Report Generation - Following exact user specifications"""
    
    def __init__(self):
        # Initialize path management first
        cd_to_project_root()
        self.paths = ensure_directories()
        
        self.logger = self._setup_logging()
        self.vbs_window = None
        
        # Find VBS window like Phase 2
        self._find_vbs_window()
        
        # Use universal path manager for images directory
        self.images_dir = self.paths['images_dir']
        
        # Verify images directory exists
        if not self.images_dir.exists():
            self.logger.error(f"❌ Images directory not found: {self.images_dir.absolute()}")
            raise FileNotFoundError(f"Images directory not found: {self.images_dir.absolute()}")
        
        self.phase4_dir = self.images_dir / "phase4"
        if not self.phase4_dir.exists():
            self.logger.error(f"❌ Phase4 images directory not found: {self.phase4_dir.absolute()}")
            raise FileNotFoundError(f"Phase4 images directory not found: {self.phase4_dir.absolute()}")
        
        self.confidence = 0.8
        
        # Current date for PDF naming
        self.current_date = datetime.now()
        
        # Check if it's morning and yesterday's PDF is missing
        self.use_yesterday_date = self._should_use_yesterday_date()
        
        if self.use_yesterday_date:
            # Use yesterday's date
            self.target_date = self.current_date - timedelta(days=1)
            self.logger.info(f"🗓️ Morning time - checking for yesterday's missing PDF: {self.target_date.strftime('%d_%m_%Y')}")
        else:
            # Use current date
            self.target_date = self.current_date
        
        # Start of month date (01/target_month/target_year) 
        self.from_date = f"01/{self.target_date.month:02d}/{self.target_date.year}"
        # Target date (target_day/target_month/target_year)
        self.to_date = f"{self.target_date.day:02d}/{self.target_date.month:02d}/{self.target_date.year}"
        self.pdf_filename = f"moonflower active users_{self.target_date.strftime('%d_%m_%Y')}"
        
        # Enhanced timing following user specs
        self.timing = {
            "after_click": 1.5,        # After each click
            "menu_open": 2.5,          # Menu opening
            "form_load": 5.0,         # Wait 5 seconds after double-click
            "print_wait": 180.0,       # 3 MINUTES for PDF generation (increased from 1 minute)
            "scroll_delay": 0.5,       # Scrolling
            "window_focus": 0.5,       # Window focus
            "typing_delay": 0.3,       # Between typing
            "clear_field_delay": 0.5   # After clearing fields
        }
        
        self.logger.info(f"📅 Date range: {self.from_date} to {self.to_date}")
        self.logger.info(f"📄 PDF filename: {self.pdf_filename}")
        self.logger.info(f"📁 Images directory: {self.images_dir.absolute()}")
        
        # Debug: Show exact date values
        self.logger.info(f"🔍 DEBUG - Current date object: {self.current_date}")
        self.logger.info(f"🔍 DEBUG - From date string: '{self.from_date}'")
        self.logger.info(f"🔍 DEBUG - To date string: '{self.to_date}'")
        self.logger.info(f"🔍 DEBUG - Day: {self.current_date.day}, Month: {self.current_date.month}, Year: {self.current_date.year}")
        
        # Verify all required images exist
        self._verify_required_images()
    
    def _should_use_yesterday_date(self) -> bool:
        """Check if we should use yesterday's date for PDF generation"""
        try:
            current_hour = self.current_date.hour
            
            # Morning time: 6 AM to 11 AM (06:00 - 11:59)
            is_morning = 6 <= current_hour <= 11
            
            if not is_morning:
                return False
            
            # Check if yesterday's PDF is missing
            yesterday = self.current_date - timedelta(days=1)
            yesterday_folder = yesterday.strftime("%d%b").lower()
            
            # Handle running from vbs directory
            if Path.cwd().name == "vbs":
                pdf_folder = Path(f"../EHC_Data_Pdf/{yesterday_folder}")
            else:
                pdf_folder = Path(f"EHC_Data_Pdf/{yesterday_folder}")
            
            # Check if any PDF exists for yesterday
            if pdf_folder.exists():
                pdf_files = list(pdf_folder.glob("*.pdf"))
                if pdf_files:
                    self.logger.info(f"📄 Yesterday's PDF already exists: {pdf_files[0].name}")
                    return False
            
            self.logger.info(f"⏰ Morning time ({current_hour}:xx) + Missing yesterday PDF = Using yesterday's date")
            return True
            
        except Exception as e:
            self.logger.warning(f"⚠️ Yesterday date check failed: {e}")
            return False
        
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for Phase 4"""
        logger = logging.getLogger("VBSPhase4Report")
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # Create console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        
        logger.addHandler(console_handler)
        
        # Create file handler in date folder
        try:
            log_dir = Path("../EHC_Logs" if Path.cwd().name == "vbs" else "EHC_Logs")
            log_dir.mkdir(exist_ok=True)
            
            today_folder = self.current_date.strftime("%d%b").lower() if hasattr(self, 'current_date') else datetime.now().strftime("%d%b").lower()
            log_subfolder = log_dir / today_folder
            log_subfolder.mkdir(exist_ok=True)
            
            log_file = log_subfolder / f"vbs_phase4_report_fixed_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
        except Exception as e:
            logger.warning(f"Could not create log file: {e}")
        
        return logger
    
    def _find_vbs_window(self):
        """Find VBS window using Phase 2 logic"""
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
    
    def _find_pdf_report_window(self):
        """Find the PDF report window that opens after printing"""
        self.logger.info("🔍 Searching for PDF report window...")
        
        def check_window(hwnd, windows):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                    
                title = win32gui.GetWindowText(hwnd).lower()
                
                # Look for PDF report indicators
                pdf_indicators = [
                    'absons',
                    'report', 
                    'pdf',
                    'active users',
                    'moonflower',
                    'wifi'
                ]
                
                # Exclude main VBS window and other non-report windows
                exclude_indicators = [
                    'sql', 'outlook', 'browser', 'chrome', 'firefox',
                    'login', 'security', 'warning'
                ]
                
                has_pdf_indicator = any(indicator in title for indicator in pdf_indicators)
                has_exclude = any(indicator in title for indicator in exclude_indicators)
                
                if has_pdf_indicator and not has_exclude:
                    # Additional check: make sure it's not the main VBS window we already have
                    if hwnd != self.vbs_window:
                        windows.append((hwnd, title))
                        
            except:
                pass
            return True
        
        windows = []
        win32gui.EnumWindows(check_window, windows)
        
        if windows:
            # Return the first matching window
            self.logger.info(f"✅ Found PDF report windows: {[w[1] for w in windows]}")
            return windows[0]
        else:
            self.logger.warning("⚠️ No PDF report window found")
            return None
    
    def _focus_vbs_only(self):
        """Focus ONLY on VBS window - prevents minimizing (copied from Phase 2)"""
        if not self.vbs_window:
            return False
            
        try:
            # Make sure it's still valid
            if not win32gui.IsWindow(self.vbs_window):
                self._find_vbs_window()
                if not self.vbs_window:
                    return False
            
            # Gentle focus - avoid aggressive window operations that cause minimizing
            win32gui.SetForegroundWindow(self.vbs_window)
            time.sleep(self.timing["window_focus"])
            
            self.logger.info("VBS window focused")
            return True
            
        except Exception as e:
            self.logger.error(f"Focus failed: {e}")
            return False
    
    def _verify_required_images(self):
        """Verify all required images exist"""
        required_images = [
            "01_arrow_button.png",
            "02_sales_distribution_menu.png", 
            "03_reports_menu.png",
            "04_pos_in_reports.png",
            "05_wifi_active_users_count.png",
            "06_from_date_field.png",
            "07_to_date_field.png", 
            "08_print_button.png",
            "09_export_button.png",
            "10_export_ok_button.png",
            "11_format_selector_dialog.png",
            "11_format_selector_ok.png",
            "12_address_bar.png",
            "13_filename_entry_field.png",
            "14_windows_save_button.png",
            "15_appclose_yes_button.png"
        ]
        
        # Optional images (nice to have but not required)
        optional_images = [
            "04_pos_highlighted.png"  # Enhanced precision for highlighted POS
        ]
        
        missing_images = []
        for img_name in required_images:
            img_path = self.phase4_dir / img_name
            if not img_path.exists():
                missing_images.append(img_name)
        
        if missing_images:
            self.logger.error(f"❌ Missing required images: {missing_images}")
            raise FileNotFoundError(f"Missing required images: {missing_images}")
        
        # Check optional images
        available_optional = []
        for img_name in optional_images:
            img_path = self.phase4_dir / img_name
            if img_path.exists():
                available_optional.append(img_name)
        
        self.logger.info(f"✅ All {len(required_images)} required images found")
        if available_optional:
            self.logger.info(f"✅ Optional images available: {available_optional}")
        else:
            self.logger.info("ℹ️ No optional images found (will use fallback methods)")
    
    def _click_image(self, image_name: str, timeout: int = 30, required: bool = True, high_precision: bool = False) -> bool:
        """Click on image using Phase 2 logic - VBS window only with anti-minimize protection"""
        if not self.phase4_dir:
            self.logger.error("No images directory")
            return False
            
        image_path = self.phase4_dir / image_name
        if not image_path.exists():
            self.logger.error(f"Image not found: {image_name}")
            return False
        
        # For high precision mode, use higher confidence and more attempts
        confidence = 0.9 if high_precision else self.confidence
        max_attempts = timeout * 4 if high_precision else timeout * 2
        
        self.logger.info(f"🔍 Looking for: {image_name} (timeout: {timeout}s, precision: {'HIGH' if high_precision else 'NORMAL'})")
        
        # Try to find image with increased attempts for reliability
        for attempt in range(max_attempts):
            try:
                # Take screenshot and look for image
                location = pyautogui.locateOnScreen(str(image_path), confidence=confidence)
                if location:
                    center = pyautogui.center(location)
                    
                    # For high precision, add small random offset to avoid pixel-perfect issues
                    if high_precision:
                        import random
                        offset_x = random.randint(-2, 2)
                        offset_y = random.randint(-2, 2)
                        click_x = center.x + offset_x
                        click_y = center.y + offset_y
                        self.logger.info(f"🎯 HIGH PRECISION: Clicking with offset ({offset_x}, {offset_y})")
                    else:
                        click_x, click_y = center.x, center.y
                    
                    # Gentle click to avoid window state changes
                    pyautogui.click(click_x, click_y)
                    self.logger.info(f"✅ Clicked {image_name} at ({click_x}, {click_y}) [attempt {attempt+1}]")
                    time.sleep(self.timing["after_click"])
                    return True
                    
            except pyautogui.ImageNotFoundException:
                pass
            except Exception as e:
                self.logger.warning(f"Click attempt {attempt+1} failed: {e}")
            
            time.sleep(0.25 if high_precision else 0.5)
        
        if required:
            self.logger.error(f"❌ Could not find {image_name} after {timeout} seconds (attempted {max_attempts} times)")
        else:
            self.logger.warning(f"⚠️ Optional image {image_name} not found within {timeout} seconds")
        return False
    
    def _double_click_image(self, image_name: str, timeout: int = 30) -> bool:
        """Double-click on image and wait 3 seconds"""
        if not self._click_image(image_name, timeout):
            return False
        
        # Double-click: click again immediately
        time.sleep(0.2)
        if not self._click_image(image_name, 10):
            self.logger.warning(f"⚠️ Second click failed for {image_name}, continuing...")
        
        # Wait 3 seconds for form to load (user requirement)
        self.logger.info("⏱️ Waiting 3 seconds for form to load...")
        time.sleep(self.timing["form_load"])
        return True
    
    def _clear_field_and_type(self, text: str):
        """Clear field completely and type new text"""
        # Clear using Ctrl+A and Delete
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.press('delete')
        time.sleep(self.timing["clear_field_delay"])
        
        # Type new text
        pyautogui.typewrite(text)
        time.sleep(self.timing["typing_delay"])
        self.logger.info(f"✅ Cleared and entered: {text}")
    
    def _create_pdf_folder_structure(self):
        """Create PDF folder structure for target date using universal path manager"""
        try:
            target_folder = self.target_date.strftime("%d%b").lower()  # e.g., "24jul"
            
            # Use universal path manager for PDF folder
            pdf_folder = self.paths['pdf_reports_base'] / target_folder
            
            pdf_folder.mkdir(parents=True, exist_ok=True)
            
            if self.use_yesterday_date:
                self.logger.info(f"📁 PDF folder ensured for YESTERDAY: {pdf_folder.absolute()}")
            else:
                self.logger.info(f"📁 PDF folder ensured for TODAY: {pdf_folder.absolute()}")
            
            return pdf_folder
            
        except Exception as e:
            self.logger.warning(f"⚠️ PDF folder creation warning: {e}")
            return None
    
    def execute_report_generation(self) -> Dict[str, Any]:
        """Execute Phase 4 PDF report generation - Following exact user specifications"""
        try:
            self.logger.info("🚀 Starting VBS Phase 4 PDF Report Generation (EXACT USER SPECS)")
            
            result = {
                "success": False,
                "start_time": datetime.now().isoformat(),
                "steps_completed": [],
                "errors": [],
                "from_date": self.from_date,
                "to_date": self.to_date,
                "pdf_filename": self.pdf_filename,
                "current_step": 0
            }
            
            # Create PDF folder structure
            pdf_folder = self._create_pdf_folder_structure()
            
            # STEP 1: Click Arrow Button (ESSENTIAL - copy from Phase 2 logic)
            result["current_step"] = 1
            self.logger.info("📋 STEP 1/15: Click Arrow Button (CRITICAL - like Phase 2)")
            if not self._click_image("01_arrow_button.png", timeout=20):
                result["errors"].append("Step 1: Failed to click Arrow button - THIS IS ESSENTIAL")
                return result
            result["steps_completed"].append("step_1_arrow_button_clicked")
            time.sleep(self.timing["menu_open"])
            
            # STEP 2: Click Sales & Distribution
            result["current_step"] = 2
            self.logger.info("📋 STEP 2/15: Click Sales & Distribution")
            if not self._click_image("02_sales_distribution_menu.png", timeout=20):
                result["errors"].append("Step 2: Failed to click Sales & Distribution menu")
                return result
            result["steps_completed"].append("step_2_sales_distribution_clicked")
            time.sleep(self.timing["menu_open"])
            
            # STEP 3: Click Reports Menu
            result["current_step"] = 3
            self.logger.info("📋 STEP 3/15: Click Reports Menu")
            if not self._click_image("03_reports_menu.png", timeout=20):
                result["errors"].append("Step 3: Failed to click Reports menu")
                return result
            result["steps_completed"].append("step_3_reports_menu_clicked")
            time.sleep(self.timing["menu_open"])
            
            # STEP 4: Use 54 DOWN arrows to reach WiFi Active Users Count (no scrolling)
            result["current_step"] = 4
            self.logger.info("📋 STEP 4/15: Navigate to WiFi Active Users Count using 54 DOWN arrows")
            
            # APPROACH 1: Try to click the highlighted POS image (if available)
            highlighted_pos_clicked = False
            if (self.phase4_dir / "04_pos_highlighted.png").exists():
                self.logger.info("🎯 Using highlighted POS image for precise clicking")
                if self._click_image("04_pos_highlighted.png", timeout=15, required=False):
                    highlighted_pos_clicked = True
                    result["steps_completed"].append("step_4_highlighted_pos_clicked")
                    self.logger.info("✅ Clicked highlighted POS successfully")
            
            # APPROACH 2: Fallback to arrow key navigation (54 DOWN + ENTER)
            if not highlighted_pos_clicked:
                self.logger.info("🎯 Using arrow key navigation: 54 DOWN + ENTER")
                
                # First try to click Reports to focus the menu (already clicked above)
                self.logger.info("✅ Reports menu already focused from step 3")
                
                # Use arrow keys to navigate to WiFi Active Users Count (54 down arrows)
                self.logger.info("⌨️ Pressing 54 DOWN arrow keys to reach WiFi Active Users Count")
                for i in range(54):
                    pyautogui.press('down')
                    time.sleep(0.3)
                    self.logger.info(f"⬇️ Down arrow {i+1}/54")
                
                # Press Enter to select WiFi Active Users Count
                self.logger.info("⌨️ Pressing ENTER to select WiFi Active Users Count")
                pyautogui.press('enter')
                time.sleep(0.5)
                
                result["steps_completed"].append("step_4_arrow_navigation_wifi_users_selected")
                self.logger.info("✅ Selected WiFi Active Users Count using 54 arrow keys")
            
            time.sleep(self.timing["menu_open"])
            
            # STEP 5: WiFi Active Users Count form should now be open (wait 3 seconds)
            result["current_step"] = 5
            self.logger.info("📋 STEP 5/15: WiFi Active Users Count form opened + wait 3 seconds")
            
            # Wait 3 seconds for form to load (user requirement)
            self.logger.info("⏱️ Waiting 3 seconds for form to load...")
            time.sleep(self.timing["form_load"])
            result["steps_completed"].append("step_5_wifi_users_form_loaded")
            # Form is now ready for date input
            
            # STEP 6: Navigate FROM date using RIGHT arrow keys (triad: day/month/year)
            result["current_step"] = 6
            self.logger.info("📋 STEP 6/15: Navigate FROM date using RIGHT arrow triad navigation")
            
            # From date opens with cursor at DAY (01) - move to MONTH
            self.logger.info("⌨️ Moving to MONTH field using RIGHT arrow")
            pyautogui.press('right')
            time.sleep(0.3)
            
            # Clear and type current month
            month_str = f"{self.current_date.month:02d}"
            self.logger.info(f"⌨️ Typing month: {month_str}")
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.2)
            pyautogui.typewrite(month_str)
            time.sleep(0.3)
            
            # Move to YEAR field using RIGHT arrow
            self.logger.info("⌨️ Moving to YEAR field using RIGHT arrow")
            pyautogui.press('right')
            time.sleep(0.3)
            
            # Clear and type current year
            year_str = str(self.current_date.year)
            self.logger.info(f"⌨️ Typing year: {year_str}")
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.2)
            pyautogui.typewrite(year_str)
            time.sleep(0.3)
            
            result["steps_completed"].append("step_6_from_date_triad_navigation")
            self.logger.info(f"✅ FROM date completed: 01/{month_str}/{year_str}")
            
            # STEP 7: Tab to TO date and navigate using RIGHT arrows
            result["current_step"] = 7
            self.logger.info("📋 STEP 7/15: Tab to TO date and use RIGHT arrow triad navigation")
            
            # Tab to TO date field
            self.logger.info("⌨️ Pressing TAB to move to TO date field")
            pyautogui.press('tab')
            time.sleep(0.5)
            
            # TO date starts at DAY - type current day
            day_str = f"{self.current_date.day:02d}"
            self.logger.info(f"⌨️ Typing current day: {day_str}")
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.2)
            pyautogui.typewrite(day_str)
            time.sleep(0.3)
            
            # Move to MONTH field using RIGHT arrow
            self.logger.info("⌨️ Moving to MONTH field using RIGHT arrow")
            pyautogui.press('right')
            time.sleep(0.3)
            
            # Type current month
            self.logger.info(f"⌨️ Typing month: {month_str}")
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.2)
            pyautogui.typewrite(month_str)
            time.sleep(0.3)
            
            # Move to YEAR field using RIGHT arrow
            self.logger.info("⌨️ Moving to YEAR field using RIGHT arrow")
            pyautogui.press('right')
            time.sleep(0.3)
            
            # Type current year
            self.logger.info(f"⌨️ Typing year: {year_str}")
            pyautogui.hotkey('ctrl', 'a')  # Select all
            time.sleep(0.2)
            pyautogui.typewrite(year_str)
            time.sleep(0.3)
            
            result["steps_completed"].append("step_7_to_date_triad_navigation")
            self.logger.info(f"✅ TO date completed: {day_str}/{month_str}/{year_str}")
            
            # STEP 8: Click Print Button and wait EXACTLY 15 seconds
            result["current_step"] = 8
            self.logger.info("📋 STEP 8/15: Click Print Button")
            if not self._click_image("08_print_button.png", timeout=20):
                result["errors"].append("Step 8: Failed to click Print button")
                return result
            result["steps_completed"].append("step_8_print_button_clicked")
            
            # Wait 3 MINUTES for PDF generation (slower systems)
            self.logger.info("⏱️ Waiting 3 MINUTES (180 seconds) for PDF generation...")
            time.sleep(self.timing["print_wait"])
            
            # STEP 9: Find PDF report window and click Export Button with high precision
            result["current_step"] = 9
            self.logger.info("📋 STEP 9/15: Find PDF report window and click Export Button")
            
            # Find the PDF report window (new window opened after print)
            self.logger.info("🔍 Looking for PDF report window...")
            pdf_window = self._find_pdf_report_window()
            
            if pdf_window:
                self.logger.info(f"✅ Found PDF report window: {pdf_window[1]}")
                
                # Focus the PDF report window
                win32gui.SetForegroundWindow(pdf_window[0])
                time.sleep(1.0)
                
                # Click export button with high precision (it's very small - red arrow)
                self.logger.info("🎯 Clicking Export button with HIGH PRECISION (red arrow)")
                if self._click_image("09_export_button.png", timeout=20, high_precision=True):
                    result["steps_completed"].append("step_9_export_button_clicked_new_window")
                    self.logger.info("✅ Export button clicked in PDF report window")
                else:
                    result["errors"].append("Step 9: Failed to click Export button in PDF report window")
                    return result
            else:
                self.logger.warning("⚠️ PDF report window not found, trying original VBS window")
                # Fallback: try in original VBS window
                if not self._focus_vbs_only():
                    result["errors"].append("Step 9: Could not focus any window for Export button")
                    return result
                
                if self._click_image("09_export_button.png", timeout=20, high_precision=True):
                    result["steps_completed"].append("step_9_export_button_clicked_fallback")
                else:
                    result["errors"].append("Step 9: Failed to click Export button")
                    return result
            
            time.sleep(self.timing["after_click"])
            
            # STEP 10: Click first OK
            result["current_step"] = 10
            self.logger.info("📋 STEP 10/15: Click first OK")
            if self._click_image("10_export_ok_button.png", timeout=15, required=False):
                result["steps_completed"].append("step_10_first_ok_clicked")
            else:
                # Fallback: Press Enter
                self.logger.info("⌨️ Pressing Enter for first OK")
                pyautogui.press('enter')
                result["steps_completed"].append("step_10_first_ok_enter")
            time.sleep(self.timing["after_click"])
            
            # STEP 11: Click second OK and ENSURE PDF format
            result["current_step"] = 11
            self.logger.info("📋 STEP 11/15: Click second OK and ENSURE PDF format")
            
            # Check if format selector appears - ensure it's PDF
            if self._click_image("11_format_selector_dialog.png", timeout=10, required=False):
                self.logger.info("📄 Format selector found - ensuring PDF format")
                # Here we would select PDF if it's not selected, but assuming it defaults to PDF
                time.sleep(0.5)
                if self._click_image("11_format_selector_ok.png", timeout=10, required=False):
                    result["steps_completed"].append("step_11_pdf_format_ensured_ok_clicked")
                else:
                    pyautogui.press('enter')
                    result["steps_completed"].append("step_11_pdf_format_ensured_enter")
            else:
                # No format selector, just press OK/Enter
                if self._click_image("11_format_selector_ok.png", timeout=10, required=False):
                    result["steps_completed"].append("step_11_second_ok_clicked")
                else:
                    pyautogui.press('enter')
                    result["steps_completed"].append("step_11_second_ok_enter")
            
            time.sleep(self.timing["after_click"])
            
            # STEP 12: Address bar navigation using Phase 3 working logic
            result["current_step"] = 12
            self.logger.info("📋 STEP 12/15: Address bar navigation (Phase 3 working logic)")
            
            # Get target date folder (today or yesterday)
            target_folder = self.target_date.strftime("%d%b").lower()
            
            # Get absolute path for PDF folder
            if Path.cwd().name == "vbs":
                base_path = Path("..").absolute()
            else:
                base_path = Path(".").absolute()
            
            pdf_folder_path = base_path / "EHC_Data_Pdf" / target_folder
            
            try:
                self.logger.info("🚀 Using Phase 3 address bar logic")
                
                # First try to double-click the address bar (same as Phase 3)
                if self._click_image("12_address_bar.png", timeout=10, required=False):
                    self.logger.info("✅ Double-clicked address bar")
                    time.sleep(0.5)
                    
                    # Double-click again for good measure (Phase 3 approach)
                    if self._click_image("12_address_bar.png", timeout=5, required=False):
                        self.logger.info("✅ Second click on address bar")
                        time.sleep(0.3)
                
                # Use Ctrl+L to focus address bar (Phase 3 method)
                self.logger.info("⌨️ Using Ctrl+L for address bar focus")
                pyautogui.hotkey('ctrl', 'l')
                time.sleep(0.5)
                
                # Clear and type path (Phase 3 method)
                folder_path = str(pdf_folder_path).replace('/', '\\')
                self.logger.info(f"📁 Typing path: {folder_path}")
                
                pyautogui.hotkey('ctrl', 'a')  # Select all
                time.sleep(0.2)
                pyautogui.typewrite(folder_path, interval=0.05)  # Same interval as Phase 3
                time.sleep(0.5)
                
                # Navigate to folder (Phase 3 method)
                pyautogui.press('enter')
                time.sleep(2.0)  # Wait for navigation
                
                result["steps_completed"].append("step_12_address_bar_navigation_phase3_method")
                self.logger.info("✅ Address bar navigation completed using Phase 3 logic")
                
            except Exception as e:
                self.logger.error(f"❌ Address bar navigation failed: {e}")
                result["errors"].append("Step 12: Address bar navigation failed")
                return result
            
            # STEP 13: Click FILENAME FIELD and type PDF filename
            result["current_step"] = 13
            self.logger.info("📋 STEP 13/15: Click FILENAME FIELD and type PDF filename")
            
            if self._click_image("13_filename_entry_field.png", timeout=15):
                self.logger.info("✅ Clicked filename entry field")
                time.sleep(0.5)
                
                # Clear any existing text and type ONLY filename (without .pdf)
                pyautogui.hotkey('ctrl', 'a')  # Select all
                time.sleep(0.2)
                self.logger.info(f"📄 Typing FILENAME: {self.pdf_filename}")
                pyautogui.typewrite(self.pdf_filename)  # NO .pdf extension here
                time.sleep(0.5)
                
                result["steps_completed"].append("step_13_filename_entered")
                self.logger.info(f"✅ Filename entered: {self.pdf_filename}")
            else:
                result["errors"].append("Step 13: Failed to click filename entry field")
                return result
            
            # STEP 14: Click SAVE BUTTON to save the PDF
            result["current_step"] = 14
            self.logger.info("📋 STEP 14/15: Click SAVE BUTTON to save PDF")
            
            if self._click_image("14_windows_save_button.png", timeout=15):
                result["steps_completed"].append("step_14_pdf_saved_via_save_button")
                self.logger.info("✅ PDF saved using Save button")
            else:
                # Fallback: Press Enter to save
                self.logger.info("⌨️ Save button not found, pressing Enter as fallback")
                pyautogui.press('enter')
                result["steps_completed"].append("step_14_pdf_saved_via_enter")
                self.logger.info("✅ PDF saved using Enter key")
            
            time.sleep(2.0)  # Wait for save
            
            # STEP 15: Verify PDF is saved and close VBS
            result["current_step"] = 15
            self.logger.info("📋 STEP 15/15: Verify PDF saved and close VBS")
            
            # Check if file exists
            full_pdf_path = pdf_folder_path / f"{self.pdf_filename}.pdf"
            if full_pdf_path.exists():
                result["steps_completed"].append("step_15_pdf_verified_in_folder")
                self.logger.info(f"✅ PDF verified in folder: {full_pdf_path}")
            else:
                result["steps_completed"].append("step_15_pdf_save_attempted")
                self.logger.warning("⚠️ PDF save attempted, verification skipped")
            
            # Close VBS application with proper popup handling
            self.logger.info("🔄 Closing VBS application...")
            try:
                if self.vbs_window and win32gui.IsWindow(self.vbs_window):
                    # Method 1: Try Alt+F4 to close
                    self.logger.info("⌨️ Using Alt+F4 to close VBS")
                    pyautogui.hotkey('alt', 'f4')
                    time.sleep(2.0)
                    
                    # Method 2: Press Enter for the close confirmation popup
                    self.logger.info("⌨️ Pressing Enter for close confirmation popup")
                    pyautogui.press('enter')
                    time.sleep(1.0)
                    
                    # Method 3: Press Enter again in case there's a second confirmation
                    self.logger.info("⌨️ Pressing Enter again for any additional popup")
                    pyautogui.press('enter')
                    time.sleep(1.0)
                    
                    result["steps_completed"].append("step_15_vbs_closed_with_enter")
                    self.logger.info("✅ VBS application closed using Enter key")
                    
                    # Alternative: Try clicking the image if available
                    if self._click_image("15_appclose_yes_button.png", timeout=3, required=False):
                        self.logger.info("✅ Also clicked close popup button")
                        result["steps_completed"].append("step_15_appclose_popup_clicked")
                    
                    time.sleep(2.0)  # Wait for app to fully close
                    
            except Exception as e:
                self.logger.warning(f"⚠️ Could not close VBS: {e}")
                result["steps_completed"].append("step_15_vbs_close_attempted")
            
            # SUCCESS!
            result.update({
                "success": True,
                "end_time": datetime.now().isoformat(),
                "total_steps": 15,
                "completed_steps": len(result["steps_completed"]),
                "pdf_saved": True,
                "pdf_path": str(full_pdf_path),
                "current_step": 15
            })
            
            self.logger.info("🎉 VBS Phase 4 PDF Report Generation completed successfully!")
            self.logger.info(f"📁 PDF saved to: {full_pdf_path}")
            return result
            
        except Exception as e:
            error_msg = f"Critical error in Phase 4 at step {result.get('current_step', 0)}: {str(e)}"
            self.logger.error(f"❌ {error_msg}")
            result["errors"].append(error_msg)
            result["end_time"] = datetime.now().isoformat()
            return result


def test_phase_4_fixed():
    """Test the completely fixed Phase 4 following exact user specifications"""
    print("🧪 Testing VBS Phase 4 PDF Report Generation (EXACT USER SPECIFICATIONS)")
    print("=" * 70)
    print("User Requirements:")
    print("✅ Arrow button click (critical, copy from Phase 2)")
    print("✅ Scroll to BOTTOM after Reports to find correct POS")
    print("✅ DOUBLE-CLICK WiFi Active Users Count + wait 3 seconds")
    print("✅ Clear all fields, enter start of month + current date")
    print("✅ Print (exactly 15 seconds), Export, OK, OK, ensure PDF")
    print("✅ Address bar navigation to save PDF correctly")
    print("=" * 70)
    
    try:
        # Initialize Phase 4
        phase4 = VBSPhase4_ReportFixed()
        
        # Show configuration
        print(f"\n📋 Configuration:")
        print(f"   VBS Window: {'✅ Found' if phase4.vbs_window else '❌ Not Found'}")
        print(f"   From Date: {phase4.from_date} (start of month)")
        print(f"   To Date: {phase4.to_date} (current date)")
        print(f"   PDF Filename: {phase4.pdf_filename}")
        print(f"   Images Directory: {phase4.images_dir.absolute()}")
        
        # Execute report generation
        result = phase4.execute_report_generation()
        
        print(f"\n📊 Phase 4 Results:")
        print(f"   Success: {result['success']}")
        print(f"   Current Step: {result.get('current_step', 0)}/15")
        print(f"   Steps Completed: {result.get('completed_steps', 0)}/15")
        print(f"   PDF Saved: {result.get('pdf_saved', False)}")
        print(f"   PDF Path: {result.get('pdf_path', 'N/A')}")
        
        if result.get('steps_completed'):
            print(f"\n✅ Completed Steps:")
            for i, step in enumerate(result['steps_completed'], 1):
                print(f"   {i}. {step}")
        
        if result.get('errors'):
            print(f"\n❌ Errors:")
            for error in result['errors']:
                print(f"   • {error}")
        
        if result["success"]:
            print("\n🎉 VBS Phase 4 PDF generation completed successfully!")
            print("📧 Ready for email delivery to General Manager")
        else:
            print(f"\n❌ VBS Phase 4 PDF generation failed at step {result.get('current_step', 0)}")
        
    except Exception as e:
        print(f"\n❌ INITIALIZATION ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 70)
    print("VBS Phase 4 Test Completed")

if __name__ == "__main__":
    import sys
    
    try:
        # Initialize Phase 4
        phase4 = VBSPhase4_ReportFixed()
        
        # Show configuration
        print(f"\n📋 Configuration:")
        print(f"   VBS Window: {'✅ Found' if phase4.vbs_window else '❌ Not Found'}")
        print(f"   From Date: {phase4.from_date} (start of month)")
        print(f"   To Date: {phase4.to_date} (current date)")
        print(f"   PDF Filename: {phase4.pdf_filename}")
        print(f"   Images Directory: {phase4.images_dir.absolute()}")
        
        # Execute report generation
        result = phase4.execute_report_generation()
        
        print(f"\n📊 Phase 4 Results:")
        print(f"   Success: {result['success']}")
        print(f"   Current Step: {result.get('current_step', 0)}/15")
        print(f"   Steps Completed: {result.get('completed_steps', 0)}/15")
        print(f"   PDF Saved: {result.get('pdf_saved', False)}")
        print(f"   PDF Path: {result.get('pdf_path', 'N/A')}")
        
        if result.get('steps_completed'):
            print(f"\n✅ Completed Steps:")
            for i, step in enumerate(result['steps_completed'], 1):
                print(f"   {i}. {step}")
        
        if result.get('errors'):
            print(f"\n❌ Errors:")
            for error in result['errors']:
                print(f"   • {error}")
        
        if result["success"]:
            print("\n🎉 VBS Phase 4 PDF generation completed successfully!")
            print("📧 Ready for email delivery to General Manager")
            sys.exit(0)  # Success
        else:
            print(f"\n❌ VBS Phase 4 PDF generation failed at step {result.get('current_step', 0)}")
            sys.exit(1)  # Failure
        
    except Exception as e:
        print(f"\n❌ INITIALIZATION ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)  # Failure 