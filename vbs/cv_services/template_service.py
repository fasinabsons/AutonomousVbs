#!/usr/bin/env python3
"""
Template Matching Service for VBS Automation
Provides image template matching capabilities for UI element detection
"""

import cv2
import numpy as np
import logging
import os
import time
import json
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass
from pathlib import Path
import win32gui
import win32ui
import win32con
import pyautogui
from .config_loader import get_cv_config

@dataclass
class TemplateMatch:
    """Represents a template match result"""
    template_name: str
    confidence: float
    location: Tuple[int, int]  # Top-left corner
    center: Tuple[int, int]
    bbox: Tuple[int, int, int, int]  # (x, y, width, height)
    scale_factor: float
    method_used: str

@dataclass
class TemplateResult:
    """Template matching operation result"""
    success: bool
    matches: List[TemplateMatch]
    processing_time: float
    template_name: str
    error_message: Optional[str] = None

class TemplateService:
    """Template matching service for VBS UI elements"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        self.template_config = self.config.get_template_config()
        self.template_cache = {}
        self.template_metadata = {}
        
        # Template directories with renamed images
        self.template_paths = {
            "phase2": "Images/phase2/",
            "phase3": "Images/phase3/", 
            "phase4": "Images/phase4/"
        }
        
        # Load all templates on initialization
        self.load_all_templates()
        
        # Load existing templates from vbs/templates (legacy)
        self._load_templates()
        
        # Performance tracking
        self.stats = {
            'total_matches': 0,
            'successful_matches': 0,
            'average_processing_time': 0.0,
            'template_usage': {},
            'confidence_distribution': []
        }
        
        self.logger.info("Template Service initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for template service"""
        logger = logging.getLogger("TemplateService")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler if debug mode is enabled
            if self.config.get('debugging.debug_mode', False):
                try:
                    log_file = "EHC_Logs/template_service.log"
                    os.makedirs(os.path.dirname(log_file), exist_ok=True)
                    file_handler = logging.FileHandler(log_file, encoding='utf-8')
                    file_handler.setFormatter(formatter)
                    logger.addHandler(file_handler)
                except Exception:
                    pass
        
        return logger
    
    def _load_templates(self):
        """Load all available templates from template directory"""
        try:
            template_dir = self.template_config.get('template_directory', 'vbs/templates')
            
            if not os.path.exists(template_dir):
                self.logger.warning(f"Template directory not found: {template_dir}")
                return
            
            # Load templates from all subdirectories
            for root, dirs, files in os.walk(template_dir):
                for file in files:
                    if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                        template_path = os.path.join(root, file)
                        template_name = os.path.splitext(file)[0]
                        
                        # Load template image
                        template_img = cv2.imread(template_path, cv2.IMREAD_COLOR)
                        if template_img is not None:
                            self.template_cache[template_name] = template_img
                            
                            # Load metadata if exists
                            metadata_path = os.path.join(root, f"{template_name}.json")
                            if os.path.exists(metadata_path):
                                with open(metadata_path, 'r') as f:
                                    self.template_metadata[template_name] = json.load(f)
                            else:
                                # Create default metadata
                                self.template_metadata[template_name] = {
                                    'created_date': time.time(),
                                    'usage_count': 0,
                                    'success_rate': 0.0,
                                    'confidence_threshold': self.template_config.get('confidence_threshold', 0.8),
                                    'scale_factors': [1.0],
                                    'description': f"Template for {template_name}"
                                }
                            
                            self.logger.info(f"Loaded template: {template_name}")
                        else:
                            self.logger.warning(f"Failed to load template: {template_path}")
            
            self.logger.info(f"Loaded {len(self.template_cache)} templates from vbs/templates")
            
        except Exception as e:
            self.logger.error(f"Error loading templates: {str(e)}")
    
    def load_all_templates(self):
        """Load all template images from renamed Images directory structure"""
        try:
            total_loaded = 0
            
            for phase, path in self.template_paths.items():
                if not os.path.exists(path):
                    self.logger.warning(f"Template directory not found: {path}")
                    continue
                
                # Load all PNG files from phase directory
                template_files = [f for f in os.listdir(path) if f.lower().endswith('.png')]
                
                for template_file in template_files:
                    template_path = os.path.join(path, template_file)
                    template_name = os.path.splitext(template_file)[0]
                    
                    # Load template image
                    template_img = cv2.imread(template_path, cv2.IMREAD_COLOR)
                    if template_img is not None:
                        # Use full name as key for uniqueness across phases
                        full_template_name = f"{phase}_{template_name}"
                        self.template_cache[full_template_name] = template_img
                        
                        # Also store with short name for backward compatibility
                        self.template_cache[template_name] = template_img
                        
                        # Create metadata for renamed templates
                        self.template_metadata[full_template_name] = {
                            'phase': phase,
                            'original_filename': template_file,
                            'file_path': template_path,
                            'created_date': time.time(),
                            'usage_count': 0,
                            'success_rate': 0.0,
                            'confidence_threshold': self.template_config.get('confidence_threshold', 0.8),
                            'scale_factors': [1.0],
                            'description': self._generate_template_description(template_name)
                        }
                        
                        total_loaded += 1
                        self.logger.info(f"✅ Loaded {phase} template: {template_name}")
                    else:
                        self.logger.warning(f"❌ Failed to load template: {template_path}")
            
            self.logger.info(f"🎯 Successfully loaded {total_loaded} templates from Images directory")
            
            # Log template inventory
            self._log_template_inventory()
            
        except Exception as e:
            self.logger.error(f"❌ Error loading templates from Images directory: {str(e)}")
    
    def _generate_template_description(self, template_name: str) -> str:
        """Generate descriptive text for template based on its name"""
        descriptions = {
            # Phase 2 templates
            "01_arrow_button": "Arrow button for menu navigation",
            "02_sales_distribution_menu": "Sales and Distribution menu item",
            "03_pos_menu": "POS menu item",
            "04_wifi_user_registration": "WiFi User Registration option",
            "05_new_button": "New button for creating entries",
            "06_credit_radio_button": "Credit radio button selection",
            
            # Phase 3 templates
            "01_import_ehc_checkbox": "Import EHC users checkbox",
            "02_three_dots_button": "Three dots file selection button",
            "03_dropdown_arrow": "Dropdown arrow for sheet selection",
            "04_sheet_selection_option": "Sheet selection option in dropdown",
            "04_sheet_selector_unselected": "Sheet selector in unselected state",
            "05_import_button": "Import button to start import process",
            "06_import_success_popup": "Import success confirmation popup",
            "06_import_ok_button": "OK button for import confirmation",
            "07_ehc_user_detail_header": "EHC user detail table header",
            "08_update_button": "Update button to start update process",
            
            # Phase 4 templates
            "01_arrow_button": "Arrow button for menu navigation",
            "02_sales_distribution_menu": "Sales and Distribution menu item",
            "03_reports_menu": "Reports menu item",
            "04_pos_in_reports": "POS option within reports section",
            "05_wifi_active_users_count": "WiFi Active Users Count report option",
            "06_from_date_field": "From Date input field",
            "07_to_date_field": "To Date input field",
            "08_print_button": "Print button for PDF generation",
            "09_export_button": "Export/Download button for PDF",
            "10_export_ok_button": "Export confirmation OK button",
            "11_format_selector_dialog": "Format selector dialog",
            "11_format_selector_ok": "Format selector OK button",
            "12_filename_entry_field": "Filename entry field",
            "13_windows_nav_arrow": "Windows navigation arrow",
            "14_windows_file_icon": "Windows file icon",
            "15_windows_save_button": "Windows save button"
        }
        
        return descriptions.get(template_name, f"Template for {template_name.replace('_', ' ')}")
    
    def _log_template_inventory(self):
        """Log complete template inventory for debugging"""
        self.logger.info("📋 Template Inventory:")
        
        for phase in ["phase2", "phase3", "phase4"]:
            phase_templates = [name for name in self.template_cache.keys() if name.startswith(phase)]
            if phase_templates:
                self.logger.info(f"   {phase.upper()}: {len(phase_templates)} templates")
                for template in sorted(phase_templates):
                    short_name = template.replace(f"{phase}_", "")
                    self.logger.info(f"      - {short_name}")
        
        legacy_templates = [name for name in self.template_cache.keys() 
                          if not any(name.startswith(phase) for phase in ["phase2", "phase3", "phase4"])]
        if legacy_templates:
            self.logger.info(f"   LEGACY: {len(legacy_templates)} templates")
    
    def get_template_by_name(self, template_name: str, phase: str = None) -> Optional[np.ndarray]:
        """
        Get template by name, with optional phase specification
        Supports both new naming convention and legacy names
        """
        # Try phase-specific name first
        if phase:
            full_name = f"{phase}_{template_name}"
            if full_name in self.template_cache:
                return self.template_cache[full_name]
        
        # Try direct name lookup
        if template_name in self.template_cache:
            return self.template_cache[template_name]
        
        # Try to find by partial match
        for cached_name in self.template_cache.keys():
            if template_name in cached_name:
                return self.template_cache[cached_name]
        
        self.logger.warning(f"⚠️ Template not found: {template_name}")
        return None
    
    def list_available_templates(self, phase: str = None) -> List[str]:
        """List all available templates, optionally filtered by phase"""
        if phase:
            return [name for name in self.template_cache.keys() if name.startswith(phase)]
        else:
            return list(self.template_cache.keys())
    
    def capture_screenshot(self, window_title: str = None) -> Optional[np.ndarray]:
        """Capture screenshot of specified window or entire screen"""
        try:
            if window_title:
                # Find window by title
                hwnd = win32gui.FindWindow(None, window_title)
                if hwnd == 0:
                    self.logger.warning(f"Window not found: {window_title}")
                    return None
                
                # Get window dimensions
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                width = right - left
                height = bottom - top
                
                # Capture window
                hwndDC = win32gui.GetWindowDC(hwnd)
                mfcDC = win32ui.CreateDCFromHandle(hwndDC)
                saveDC = mfcDC.CreateCompatibleDC()
                
                saveBitMap = win32ui.CreateBitmap()
                saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
                saveDC.SelectObject(saveBitMap)
                
                # Copy window content
                saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
                
                # Convert to numpy array
                bmpinfo = saveBitMap.GetInfo()
                bmpstr = saveBitMap.GetBitmapBits(True)
                
                img = np.frombuffer(bmpstr, dtype='uint8')
                img.shape = (height, width, 4)
                img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                
                # Cleanup
                win32gui.DeleteObject(saveBitMap.GetHandle())
                saveDC.DeleteDC()
                mfcDC.DeleteDC()
                win32gui.ReleaseDC(hwnd, hwndDC)
                
                return img
            else:
                # Capture entire screen
                import pyautogui
                screenshot = pyautogui.screenshot()
                return cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
                
        except Exception as e:
            self.logger.error(f"Screenshot capture failed: {str(e)}")
            return None
    
    def preprocess_image(self, image: np.ndarray, enhance: bool = True) -> np.ndarray:
        """Preprocess image for better template matching"""
        try:
            processed = image.copy()
            
            if enhance:
                # Convert to grayscale for processing
                if len(processed.shape) == 3:
                    gray = cv2.cvtColor(processed, cv2.COLOR_BGR2GRAY)
                else:
                    gray = processed
                
                # Apply CLAHE (Contrast Limited Adaptive Histogram Equalization)
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                enhanced = clahe.apply(gray)
                
                # Apply Gaussian blur to reduce noise
                blurred = cv2.GaussianBlur(enhanced, (3, 3), 0)
                
                # Convert back to BGR if needed
                if len(image.shape) == 3:
                    processed = cv2.cvtColor(blurred, cv2.COLOR_GRAY2BGR)
                else:
                    processed = blurred
            
            return processed
            
        except Exception as e:
            self.logger.error(f"Image preprocessing failed: {str(e)}")
            return image
    
    def match_template(self, template_name: str, screenshot: np.ndarray = None, 
                      window_title: str = None, confidence_threshold: float = None,
                      scale_factors: List[float] = None, max_matches: int = 1) -> TemplateResult:
        """
        Match template against screenshot with multi-scale support
        
        Args:
            template_name: Name of template to match
            screenshot: Screenshot to search in (if None, captures new one)
            window_title: Window to capture if screenshot is None
            confidence_threshold: Minimum confidence for match (overrides config)
            scale_factors: List of scale factors to try (overrides metadata)
            max_matches: Maximum number of matches to return
        """
        start_time = time.time()
        
        try:
            # Get template
            if template_name not in self.template_cache:
                return TemplateResult(
                    success=False,
                    matches=[],
                    processing_time=time.time() - start_time,
                    template_name=template_name,
                    error_message=f"Template not found: {template_name}"
                )
            
            template = self.template_cache[template_name]
            metadata = self.template_metadata.get(template_name, {})
            
            # Get screenshot
            if screenshot is None:
                screenshot = self.capture_screenshot(window_title)
                if screenshot is None:
                    return TemplateResult(
                        success=False,
                        matches=[],
                        processing_time=time.time() - start_time,
                        template_name=template_name,
                        error_message="Failed to capture screenshot"
                    )
            
            # Set parameters
            if confidence_threshold is None:
                confidence_threshold = metadata.get('confidence_threshold', 
                                                  self.template_config.get('confidence_threshold', 0.8))
            
            if scale_factors is None:
                scale_factors = metadata.get('scale_factors', [1.0])
            
            # Preprocess images
            processed_screenshot = self.preprocess_image(screenshot)
            processed_template = self.preprocess_image(template)
            
            matches = []
            best_confidence = 0.0
            
            # Try different matching methods
            methods = [
                ('TM_CCOEFF_NORMED', cv2.TM_CCOEFF_NORMED),
                ('TM_CCORR_NORMED', cv2.TM_CCORR_NORMED),
                ('TM_SQDIFF_NORMED', cv2.TM_SQDIFF_NORMED)
            ]
            
            for scale in scale_factors:
                # Scale template
                if scale != 1.0:
                    h, w = processed_template.shape[:2]
                    new_h, new_w = int(h * scale), int(w * scale)
                    scaled_template = cv2.resize(processed_template, (new_w, new_h))
                else:
                    scaled_template = processed_template
                
                # Skip if template is larger than screenshot
                if (scaled_template.shape[0] > processed_screenshot.shape[0] or 
                    scaled_template.shape[1] > processed_screenshot.shape[1]):
                    continue
                
                for method_name, method in methods:
                    try:
                        # Perform template matching
                        result = cv2.matchTemplate(processed_screenshot, scaled_template, method)
                        
                        # Find matches
                        if method == cv2.TM_SQDIFF_NORMED:
                            # For SQDIFF, lower values are better
                            locations = np.where(result <= (1.0 - confidence_threshold))
                            confidences = 1.0 - result[locations]
                        else:
                            # For other methods, higher values are better
                            locations = np.where(result >= confidence_threshold)
                            confidences = result[locations]
                        
                        # Process matches
                        for i, (y, x) in enumerate(zip(locations[0], locations[1])):
                            confidence = float(confidences[i])
                            
                            if confidence > best_confidence:
                                best_confidence = confidence
                            
                            # Calculate match details
                            h, w = scaled_template.shape[:2]
                            center_x = x + w // 2
                            center_y = y + h // 2
                            
                            match = TemplateMatch(
                                template_name=template_name,
                                confidence=confidence,
                                location=(x, y),
                                center=(center_x, center_y),
                                bbox=(x, y, w, h),
                                scale_factor=scale,
                                method_used=method_name
                            )
                            
                            matches.append(match)
                    
                    except Exception as e:
                        self.logger.warning(f"Template matching failed for method {method_name}: {str(e)}")
                        continue
            
            # Sort matches by confidence and remove duplicates
            matches.sort(key=lambda m: m.confidence, reverse=True)
            
            # Remove overlapping matches (Non-Maximum Suppression)
            filtered_matches = self._apply_nms(matches, overlap_threshold=0.3)
            
            # Limit number of matches
            filtered_matches = filtered_matches[:max_matches]
            
            # Update statistics
            self._update_stats(template_name, len(filtered_matches) > 0, time.time() - start_time, best_confidence)
            
            # Save debug screenshot if enabled and matches found
            if (self.config.get('debugging.save_screenshots', False) and 
                len(filtered_matches) > 0):
                self._save_debug_screenshot(screenshot, filtered_matches, template_name)
            
            return TemplateResult(
                success=len(filtered_matches) > 0,
                matches=filtered_matches,
                processing_time=time.time() - start_time,
                template_name=template_name
            )
            
        except Exception as e:
            self.logger.error(f"Template matching error for {template_name}: {str(e)}")
            return TemplateResult(
                success=False,
                matches=[],
                processing_time=time.time() - start_time,
                template_name=template_name,
                error_message=str(e)
            )
    
    def _apply_nms(self, matches: List[TemplateMatch], overlap_threshold: float = 0.3) -> List[TemplateMatch]:
        """Apply Non-Maximum Suppression to remove overlapping matches"""
        if not matches:
            return matches
        
        # Convert to format suitable for NMS
        boxes = []
        scores = []
        
        for match in matches:
            x, y, w, h = match.bbox
            boxes.append([x, y, x + w, y + h])
            scores.append(match.confidence)
        
        boxes = np.array(boxes, dtype=np.float32)
        scores = np.array(scores, dtype=np.float32)
        
        # Apply NMS
        indices = cv2.dnn.NMSBoxes(boxes.tolist(), scores.tolist(), 
                                  score_threshold=0.0, nms_threshold=overlap_threshold)
        
        if len(indices) > 0:
            indices = indices.flatten()
            return [matches[i] for i in indices]
        else:
            return []
    
    def _update_stats(self, template_name: str, success: bool, processing_time: float, confidence: float):
        """Update performance statistics"""
        self.stats['total_matches'] += 1
        if success:
            self.stats['successful_matches'] += 1
        
        # Update average processing time
        current_avg = self.stats['average_processing_time']
        total = self.stats['total_matches']
        self.stats['average_processing_time'] = ((current_avg * (total - 1)) + processing_time) / total
        
        # Update template usage
        if template_name not in self.stats['template_usage']:
            self.stats['template_usage'][template_name] = {'count': 0, 'success_count': 0}
        
        self.stats['template_usage'][template_name]['count'] += 1
        if success:
            self.stats['template_usage'][template_name]['success_count'] += 1
        
        # Update confidence distribution
        if success:
            self.stats['confidence_distribution'].append(confidence)
            # Keep only last 100 confidence scores
            if len(self.stats['confidence_distribution']) > 100:
                self.stats['confidence_distribution'] = self.stats['confidence_distribution'][-100:]
        
        # Update template metadata
        if template_name in self.template_metadata:
            metadata = self.template_metadata[template_name]
            metadata['usage_count'] = self.stats['template_usage'][template_name]['count']
            
            success_count = self.stats['template_usage'][template_name]['success_count']
            total_count = metadata['usage_count']
            metadata['success_rate'] = success_count / total_count if total_count > 0 else 0.0
    
    def _save_debug_screenshot(self, screenshot: np.ndarray, matches: List[TemplateMatch], template_name: str):
        """Save debug screenshot with match annotations"""
        try:
            debug_img = screenshot.copy()
            
            # Draw bounding boxes for matches
            for i, match in enumerate(matches):
                x, y, w, h = match.bbox
                
                # Draw rectangle
                color = (0, 255, 0) if i == 0 else (0, 255, 255)  # Green for best match, yellow for others
                cv2.rectangle(debug_img, (x, y), (x + w, y + h), color, 2)
                
                # Draw center point
                cv2.circle(debug_img, match.center, 5, color, -1)
                
                # Add confidence text
                text = f"{match.confidence:.2f}"
                cv2.putText(debug_img, text, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 1)
            
            # Save debug image
            debug_dir = "EHC_Logs/template_debug"
            os.makedirs(debug_dir, exist_ok=True)
            
            timestamp = int(time.time())
            filename = f"{debug_dir}/{template_name}_match_{timestamp}.png"
            cv2.imwrite(filename, debug_img)
            
            self.logger.debug(f"Debug screenshot saved: {filename}")
            
        except Exception as e:
            self.logger.error(f"Failed to save debug screenshot: {str(e)}")
    
    def find_element(self, template_name: str, window_title: str = None, 
                    timeout: float = 10.0, retry_interval: float = 1.0) -> Optional[TemplateMatch]:
        """
        Find UI element using template matching with retry logic
        
        Args:
            template_name: Name of template to find
            window_title: Window to search in
            timeout: Maximum time to search (seconds)
            retry_interval: Time between retries (seconds)
        """
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            result = self.match_template(template_name, window_title=window_title)
            
            if result.success and result.matches:
                return result.matches[0]  # Return best match
            
            time.sleep(retry_interval)
        
        self.logger.warning(f"Element not found after {timeout}s: {template_name}")
        return None
    
    def click_template(self, template_name: str, window_title: str = None, 
                      timeout: float = 10.0, offset: Tuple[int, int] = (0, 0)) -> bool:
        """
        Find and click on template match
        
        Args:
            template_name: Name of template to click
            window_title: Window to search in
            timeout: Maximum time to search
            offset: Offset from center to click (x, y)
        """
        match = self.find_element(template_name, window_title, timeout)
        
        if match:
            try:
                import pyautogui
                click_x = match.center[0] + offset[0]
                click_y = match.center[1] + offset[1]
                
                pyautogui.click(click_x, click_y)
                self.logger.info(f"Clicked template {template_name} at ({click_x}, {click_y})")
                return True
                
            except Exception as e:
                self.logger.error(f"Failed to click template {template_name}: {str(e)}")
                return False
        
        return False
    
    def get_template_list(self) -> List[str]:
        """Get list of available template names"""
        return list(self.template_cache.keys())
    
    def get_template_info(self, template_name: str) -> Optional[Dict[str, Any]]:
        """Get information about a specific template"""
        if template_name not in self.template_metadata:
            return None
        
        metadata = self.template_metadata[template_name].copy()
        
        # Add runtime statistics
        if template_name in self.stats['template_usage']:
            usage = self.stats['template_usage'][template_name]
            metadata['runtime_stats'] = {
                'total_usage': usage['count'],
                'success_count': usage['success_count'],
                'current_success_rate': usage['success_count'] / usage['count'] if usage['count'] > 0 else 0.0
            }
        
        return metadata
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get overall performance statistics"""
        stats = self.stats.copy()
        
        # Calculate additional metrics
        if stats['total_matches'] > 0:
            stats['overall_success_rate'] = stats['successful_matches'] / stats['total_matches']
        else:
            stats['overall_success_rate'] = 0.0
        
        if stats['confidence_distribution']:
            stats['average_confidence'] = sum(stats['confidence_distribution']) / len(stats['confidence_distribution'])
            stats['min_confidence'] = min(stats['confidence_distribution'])
            stats['max_confidence'] = max(stats['confidence_distribution'])
        
        return stats
    
    def save_template(self, template_name: str, image: np.ndarray, 
                     metadata: Dict[str, Any] = None) -> bool:
        """
        Save new template to template directory
        
        Args:
            template_name: Name for the template
            image: Template image
            metadata: Optional metadata for template
        """
        try:
            template_dir = self.template_config.get('template_directory', 'vbs/templates')
            os.makedirs(template_dir, exist_ok=True)
            
            # Save image
            image_path = os.path.join(template_dir, f"{template_name}.png")
            cv2.imwrite(image_path, image)
            
            # Save metadata
            if metadata is None:
                metadata = {
                    'created_date': time.time(),
                    'usage_count': 0,
                    'success_rate': 0.0,
                    'confidence_threshold': self.template_config.get('confidence_threshold', 0.8),
                    'scale_factors': [1.0],
                    'description': f"Template for {template_name}"
                }
            
            metadata_path = os.path.join(template_dir, f"{template_name}.json")
            with open(metadata_path, 'w') as f:
                json.dump(metadata, f, indent=2)
            
            # Update cache
            self.template_cache[template_name] = image
            self.template_metadata[template_name] = metadata
            
            self.logger.info(f"Template saved: {template_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save template {template_name}: {str(e)}")
            return False
