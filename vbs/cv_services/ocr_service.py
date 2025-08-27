#!/usr/bin/env python3
"""
OCR Service for VBS Automation
Provides text recognition capabilities using Tesseract OCR with fallback options
"""

import cv2
import numpy as np
import pytesseract
import logging
import os
import time
from typing import List, Dict, Tuple, Optional, Any
from dataclasses import dataclass
from PIL import Image, ImageEnhance, ImageFilter
import win32gui
import win32ui
import win32con
try:
    import winrt.windows.media.ocr as winrt_ocr
    import winrt.windows.graphics.imaging as winrt_imaging
    import winrt.windows.storage.streams as winrt_streams
    WINDOWS_OCR_AVAILABLE = True
except ImportError:
    WINDOWS_OCR_AVAILABLE = False
from .config_loader import get_cv_config

@dataclass
class TextMatch:
    """Represents a detected text match"""
    text: str
    confidence: float
    bbox: Tuple[int, int, int, int]  # (x, y, width, height)
    center: Tuple[int, int]
    clickable_region: Tuple[int, int, int, int]

@dataclass
class OCRResult:
    """OCR operation result"""
    success: bool
    matches: List[TextMatch]
    processing_time: float
    method_used: str
    error_message: Optional[str] = None

class OCRService:
    """OCR service for VBS text recognition"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        self.ocr_config = self.config.get_ocr_config()
        self._setup_tesseract()
        self._setup_windows_ocr()
        
        # Performance tracking
        self.stats = {
            'total_operations': 0,
            'successful_operations': 0,
            'average_processing_time': 0.0,
            'method_success_rates': {
                'tesseract': 0,
                'windows_ocr': 0
            }
        }
        
        self.logger.info("OCR Service initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for OCR service"""
        logger = logging.getLogger("OCRService")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler if debug mode is enabled
            if self.config.get('debugging.debug_mode', False):
                try:
                    log_file = "EHC_Logs/ocr_service.log"
                    os.makedirs(os.path.dirname(log_file), exist_ok=True)
                    file_handler = logging.FileHandler(log_file, encoding='utf-8')
                    file_handler.setFormatter(formatter)
                    logger.addHandler(file_handler)
                except Exception:
                    pass
        
        return logger
    
    def _setup_tesseract(self):
        """Setup Tesseract OCR configuration"""
        try:
            # Set Tesseract path if specified
            tesseract_path = self.ocr_config.get('tesseract_path')
            if tesseract_path and os.path.exists(tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                self.logger.info(f"Tesseract path set to: {tesseract_path}")
            
            # Test Tesseract installation
            version = pytesseract.get_tesseract_version()
            self.logger.info(f"Tesseract version: {version}")
            
        except Exception as e:
            self.logger.warning(f"Tesseract setup warning: {e}")
    
    def _setup_windows_ocr(self):
        """Setup Windows OCR as fallback"""
        try:
            if WINDOWS_OCR_AVAILABLE:
                # Check if Windows OCR is available
                self.windows_ocr_engine = winrt_ocr.OcrEngine.try_create_from_language(
                    winrt_ocr.Language("en-US")
                )
                if self.windows_ocr_engine:
                    self.logger.info("Windows OCR engine initialized successfully")
                else:
                    self.logger.warning("Windows OCR engine could not be created")
            else:
                self.windows_ocr_engine = None
                self.logger.warning("Windows OCR not available - winrt packages not installed")
        except Exception as e:
            self.windows_ocr_engine = None
            self.logger.warning(f"Windows OCR setup failed: {e}")
    
    def preprocess_image_for_ocr(self, image: np.ndarray, enhance_text: bool = True) -> np.ndarray:
        """Preprocess image for better OCR accuracy"""
        try:
            # Convert to PIL Image for enhancement
            if len(image.shape) == 3:
                pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
            else:
                pil_image = Image.fromarray(image)
            
            if enhance_text:
                # Enhance contrast and sharpness
                enhancer = ImageEnhance.Contrast(pil_image)
                pil_image = enhancer.enhance(1.5)
                
                enhancer = ImageEnhance.Sharpness(pil_image)
                pil_image = enhancer.enhance(1.2)
            
            # Convert back to OpenCV format
            if pil_image.mode != 'RGB':
                pil_image = pil_image.convert('RGB')
            
            processed = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
            
            # Convert to grayscale
            if len(processed.shape) == 3:
                gray = cv2.cvtColor(processed, cv2.COLOR_BGR2GRAY)
            else:
                gray = processed
            
            # Apply Gaussian blur to reduce noise
            preprocessing = self.ocr_config.get('preprocessing', {})
            kernel_size = preprocessing.get('gaussian_blur_kernel', 3)
            if kernel_size > 0:
                gray = cv2.GaussianBlur(gray, (kernel_size, kernel_size), 0)
            
            # Apply adaptive threshold for better text contrast
            block_size = preprocessing.get('adaptive_threshold_block_size', 11)
            c_value = preprocessing.get('adaptive_threshold_c', 2)
            
            binary = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY, block_size, c_value
            )
            
            # Morphological operations to clean up text
            kernel_size = preprocessing.get('morphology_kernel_size', 2)
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
            cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
            
            return cleaned
            
        except Exception as e:
            self.logger.error(f"Image preprocessing failed: {e}")
            return image
    
    def extract_text_tesseract(self, image: np.ndarray, region: Optional[Tuple[int, int, int, int]] = None) -> OCRResult:
        """Extract text using Tesseract OCR"""
        start_time = time.time()
        
        try:
            # Extract region if specified
            if region:
                x, y, w, h = region
                roi = image[y:y+h, x:x+w]
            else:
                roi = image
            
            # Preprocess image
            processed = self.preprocess_image_for_ocr(roi)
            
            # Configure Tesseract
            config = self._get_tesseract_config()
            
            # Extract text with bounding boxes
            data = pytesseract.image_to_data(processed, config=config, output_type=pytesseract.Output.DICT)
            
            # Process results
            matches = []
            confidence_threshold = self.ocr_config.get('confidence_threshold', 0.7) * 100
            
            for i in range(len(data['text'])):
                text = data['text'][i].strip()
                confidence = float(data['conf'][i])
                
                if text and confidence >= confidence_threshold:
                    x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                    
                    # Adjust coordinates if region was specified
                    if region:
                        x += region[0]
                        y += region[1]
                    
                    # Calculate center and clickable region
                    center = (x + w // 2, y + h // 2)
                    clickable_region = (x - 5, y - 5, w + 10, h + 10)  # Add padding
                    
                    match = TextMatch(
                        text=text,
                        confidence=confidence / 100.0,
                        bbox=(x, y, w, h),
                        center=center,
                        clickable_region=clickable_region
                    )
                    matches.append(match)
            
            processing_time = time.time() - start_time
            
            # Update stats
            self.stats['total_operations'] += 1
            self.stats['successful_operations'] += 1
            self.stats['method_success_rates']['tesseract'] += 1
            
            # Update average processing time
            if self.stats['total_operations'] > 0:
                current_avg = self.stats['average_processing_time']
                self.stats['average_processing_time'] = (
                    (current_avg * (self.stats['total_operations'] - 1) + processing_time) / 
                    self.stats['total_operations']
                )
            
            return OCRResult(
                success=True,
                matches=matches,
                processing_time=processing_time,
                method_used='tesseract'
            )
            
        except Exception as e:
            processing_time = time.time() - start_time
            self.logger.error(f"Tesseract OCR failed: {e}")
            
            return OCRResult(
                success=False,
                matches=[],
                processing_time=processing_time,
                method_used='tesseract',
                error_message=str(e)
            )
    
    def _get_tesseract_config(self) -> str:
        """Get Tesseract configuration string"""
        config_parts = []
        
        # Page segmentation mode
        psm = self.ocr_config.get('page_segmentation_mode', 6)
        config_parts.append(f'--psm {psm}')
        
        # Character whitelist
        whitelist = self.ocr_config.get('character_whitelist')
        if whitelist:
            config_parts.append(f'-c tessedit_char_whitelist={whitelist}')
        
        # Language
        language = self.ocr_config.get('language', 'eng')
        config_parts.append(f'-l {language}')
        
        return ' '.join(config_parts)
    
    def extract_text_windows_ocr(self, image: np.ndarray, region: Optional[Tuple[int, int, int, int]] = None) -> OCRResult:
        """Extract text using Windows OCR API as fallback"""
        start_time = time.time()
        
        try:
            if not self.windows_ocr_engine:
                return OCRResult(
                    success=False,
                    matches=[],
                    processing_time=time.time() - start_time,
                    method_used='windows_ocr',
                    error_message="Windows OCR engine not available"
                )
            
            # Extract region if specified
            if region:
                x, y, w, h = region
                roi = image[y:y+h, x:x+w]
            else:
                roi = image
            
            # Convert to PIL Image and then to bytes
            if len(roi.shape) == 3:
                pil_image = Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
            else:
                pil_image = Image.fromarray(roi)
            
            # Convert to bytes for Windows OCR
            import io
            img_bytes = io.BytesIO()
            pil_image.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            
            # Create Windows Runtime stream
            stream = winrt_streams.InMemoryRandomAccessStream()
            writer = winrt_streams.DataWriter(stream.get_output_stream_at(0))
            writer.write_bytes(img_bytes.getvalue())
            writer.store_async().get()
            writer.close()
            
            # Create bitmap from stream
            decoder = winrt_imaging.BitmapDecoder.create_async(stream).get()
            bitmap = decoder.get_software_bitmap_async().get()
            
            # Perform OCR
            ocr_result = self.windows_ocr_engine.recognize_async(bitmap).get()
            
            # Process results
            matches = []
            confidence_threshold = self.ocr_config.get('confidence_threshold', 0.7)
            
            for line in ocr_result.lines:
                for word in line.words:
                    text = word.text.strip()
                    # Windows OCR doesn't provide confidence, so we use a default high value
                    confidence = 0.9
                    
                    if text and confidence >= confidence_threshold:
                        # Get bounding box
                        bbox_rect = word.bounding_rect
                        x, y, w, h = int(bbox_rect.x), int(bbox_rect.y), int(bbox_rect.width), int(bbox_rect.height)
                        
                        # Adjust coordinates if region was specified
                        if region:
                            x += region[0]
                            y += region[1]
                        
                        # Calculate center and clickable region
                        center = (x + w // 2, y + h // 2)
                        clickable_region = (x - 5, y - 5, w + 10, h + 10)  # Add padding
                        
                        match = TextMatch(
                            text=text,
                            confidence=confidence,
                            bbox=(x, y, w, h),
                            center=center,
                            clickable_region=clickable_region
                        )
                        matches.append(match)
            
            processing_time = time.time() - start_time
            
            # Update stats
            self.stats['method_success_rates']['windows_ocr'] += 1
            
            return OCRResult(
                success=True,
                matches=matches,
                processing_time=processing_time,
                method_used='windows_ocr'
            )
            
        except Exception as e:
            processing_time = time.time() - start_time
            self.logger.error(f"Windows OCR failed: {e}")
            
            return OCRResult(
                success=False,
                matches=[],
                processing_time=processing_time,
                method_used='windows_ocr',
                error_message=str(e)
            )
    
    def extract_text_with_fallback(self, image: np.ndarray, region: Optional[Tuple[int, int, int, int]] = None) -> OCRResult:
        """Extract text using Tesseract with Windows OCR fallback"""
        # Try Tesseract first
        result = self.extract_text_tesseract(image, region)
        
        # If Tesseract fails or has low confidence, try Windows OCR
        if not result.success or (result.matches and 
                                 sum(m.confidence for m in result.matches) / len(result.matches) < 0.6):
            self.logger.info("Tesseract failed or low confidence, trying Windows OCR fallback")
            windows_result = self.extract_text_windows_ocr(image, region)
            
            if windows_result.success and len(windows_result.matches) > len(result.matches):
                return windows_result
        
        return result
    
    def find_text(self, image: np.ndarray, search_text: str, case_sensitive: bool = False) -> List[TextMatch]:
        """Find specific text in image"""
        try:
            # Extract all text with fallback
            result = self.extract_text_with_fallback(image)
            
            if not result.success:
                return []
            
            # Filter matches by search text
            matches = []
            search_lower = search_text.lower() if not case_sensitive else search_text
            
            for match in result.matches:
                match_text = match.text if case_sensitive else match.text.lower()
                
                if search_lower in match_text:
                    matches.append(match)
            
            self.logger.info(f"Found {len(matches)} matches for text '{search_text}' using {result.method_used}")
            return matches
            
        except Exception as e:
            self.logger.error(f"Text search failed: {e}")
            return []
    
    def find_vbs_ui_elements(self, image: np.ndarray, element_names: List[str]) -> Dict[str, List[TextMatch]]:
        """Find VBS UI elements by their text labels"""
        try:
            results = {}
            
            # Common VBS UI element variations
            element_variations = {
                'new': ['New', 'NEW', 'new'],
                'update': ['Update', 'UPDATE', 'update'],
                'import': ['Import', 'IMPORT', 'import'],
                'export': ['Export', 'EXPORT', 'export'],
                'print': ['Print', 'PRINT', 'print'],
                'ok': ['OK', 'Ok', 'ok'],
                'cancel': ['Cancel', 'CANCEL', 'cancel'],
                'sales_distribution': ['Sales & Distribution', 'Sales and Distribution', 'Sales'],
                'pos': ['POS', 'pos', 'Point of Sale'],
                'wifi_registration': ['WiFi User Registration', 'WiFi Registration', 'User Registration'],
                'reports': ['Reports', 'REPORTS', 'reports']
            }
            
            for element_name in element_names:
                element_matches = []
                variations = element_variations.get(element_name.lower(), [element_name])
                
                for variation in variations:
                    matches = self.find_text(image, variation, case_sensitive=False)
                    element_matches.extend(matches)
                
                # Remove duplicates based on proximity
                element_matches = self._remove_duplicate_matches(element_matches)
                results[element_name] = element_matches
                
                self.logger.info(f"Found {len(element_matches)} matches for element '{element_name}'")
            
            return results
            
        except Exception as e:
            self.logger.error(f"VBS UI element search failed: {e}")
            return {}
    
    def _remove_duplicate_matches(self, matches: List[TextMatch], proximity_threshold: int = 20) -> List[TextMatch]:
        """Remove duplicate matches that are too close to each other"""
        if not matches:
            return matches
        
        # Sort by confidence (highest first)
        sorted_matches = sorted(matches, key=lambda m: m.confidence, reverse=True)
        filtered_matches = []
        
        for match in sorted_matches:
            is_duplicate = False
            
            for existing in filtered_matches:
                # Calculate distance between centers
                dx = abs(match.center[0] - existing.center[0])
                dy = abs(match.center[1] - existing.center[1])
                distance = (dx ** 2 + dy ** 2) ** 0.5
                
                if distance < proximity_threshold:
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                filtered_matches.append(match)
        
        return filtered_matches
    
    def capture_screenshot_region(self, window_handle: int, region: Optional[Tuple[int, int, int, int]] = None) -> Optional[np.ndarray]:
        """Capture screenshot of window or region"""
        try:
            # Get window dimensions
            window_rect = win32gui.GetWindowRect(window_handle)
            window_width = window_rect[2] - window_rect[0]
            window_height = window_rect[3] - window_rect[1]
            
            # Create device context
            hwndDC = win32gui.GetWindowDC(window_handle)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            
            # Create bitmap
            if region:
                x, y, w, h = region
                saveBitMap = win32ui.CreateBitmap()
                saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
                saveDC.SelectObject(saveBitMap)
                saveDC.BitBlt((0, 0), (w, h), mfcDC, (x, y), win32con.SRCCOPY)
            else:
                saveBitMap = win32ui.CreateBitmap()
                saveBitMap.CreateCompatibleBitmap(mfcDC, window_width, window_height)
                saveDC.SelectObject(saveBitMap)
                saveDC.BitBlt((0, 0), (window_width, window_height), mfcDC, (0, 0), win32con.SRCCOPY)
            
            # Convert to numpy array
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            
            img = np.frombuffer(bmpstr, dtype='uint8')
            img.shape = (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4)
            
            # Convert BGRA to BGR
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
            # Cleanup
            win32gui.DeleteObject(saveBitMap.GetHandle())
            saveDC.DeleteDC()
            mfcDC.DeleteDC()
            win32gui.ReleaseDC(window_handle, hwndDC)
            
            return img
            
        except Exception as e:
            self.logger.error(f"Screenshot capture failed: {e}")
            return None
    
    def save_debug_image(self, image: np.ndarray, filename: str, matches: List[TextMatch] = None):
        """Save debug image with detected text highlighted"""
        try:
            if not self.config.get('debugging.save_debug_images', False):
                return
            
            debug_dir = self.config.get('debugging.debug_image_path', 'debug_images')
            os.makedirs(debug_dir, exist_ok=True)
            
            debug_image = image.copy()
            
            # Draw bounding boxes for detected text
            if matches:
                for match in matches:
                    x, y, w, h = match.bbox
                    cv2.rectangle(debug_image, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(debug_image, f"{match.text} ({match.confidence:.2f})", 
                               (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
            
            # Save image
            debug_path = os.path.join(debug_dir, filename)
            cv2.imwrite(debug_path, debug_image)
            self.logger.debug(f"Debug image saved: {debug_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to save debug image: {e}")
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get OCR performance statistics"""
        return self.stats.copy()
    
    def reset_stats(self):
        """Reset performance statistics"""
        self.stats = {
            'total_operations': 0,
            'successful_operations': 0,
            'average_processing_time': 0.0,
            'method_success_rates': {
                'tesseract': 0,
                'windows_ocr': 0
            }
        }

if __name__ == "__main__":
    # Test OCR service
    ocr = OCRService()
    print("OCR Service initialized successfully")
    print("Configuration:", ocr.ocr_config)