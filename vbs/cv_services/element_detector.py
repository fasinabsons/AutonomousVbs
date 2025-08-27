#!/usr/bin/env python3
"""
UI Element Detection Service for VBS Automation
Provides unified element detection using multiple methods with confidence scoring
"""

import time
import logging
import os
from typing import List, Dict, Tuple, Optional, Any, Union
from dataclasses import dataclass
from enum import Enum
import cv2
import numpy as np
from .ocr_service import OCRService, TextMatch
from .template_service import TemplateService, TemplateMatch
from .config_loader import get_cv_config

class DetectionMethod(Enum):
    """Available detection methods"""
    OCR = "ocr"
    TEMPLATE = "template"
    HYBRID = "hybrid"

@dataclass
class ElementDescriptor:
    """Describes a UI element to be detected"""
    name: str
    text_patterns: Optional[List[str]] = None
    template_names: Optional[List[str]] = None
    region: Optional[Tuple[int, int, int, int]] = None
    confidence_threshold: float = 0.7
    method_preference: Optional[DetectionMethod] = None
    clickable_offset: Tuple[int, int] = (0, 0)  # Offset from detected center for clicking

@dataclass
class ElementMatch:
    """Represents a detected UI element"""
    element_name: str
    method_used: str
    confidence: float
    location: Tuple[int, int]  # Center point
    bounding_box: Tuple[int, int, int, int]  # (x, y, width, height)
    clickable_region: Tuple[int, int, int, int]  # Expanded clickable area
    detection_time: float
    raw_match: Union[TextMatch, TemplateMatch, None] = None

@dataclass
class DetectionResult:
    """Result of element detection operation"""
    success: bool
    matches: List[ElementMatch]
    total_time: float
    methods_tried: List[str]
    error_message: Optional[str] = None
    screenshot_path: Optional[str] = None

class UIElementDetector:
    """Unified UI element detection service"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        
        # Initialize detection services
        self.ocr_service = OCRService()
        self.template_service = TemplateService()
        
        # Configuration
        self.default_confidence = self.config.get('smart_automation.confidence_threshold', 0.7)
        self.region_optimization = self.config.get('performance.screenshot_region_optimization', True)
        self.cache_enabled = self.config.get('performance.cache_successful_locations', True)
        self.cache_duration = self.config.get('performance.cache_duration_seconds', 300)
        
        # Element cache for performance
        self.element_cache = {}
        
        # Performance tracking
        self.stats = {
            'total_detections': 0,
            'successful_detections': 0,
            'method_usage': {method.value: 0 for method in DetectionMethod},
            'average_detection_time': 0.0,
            'cache_hits': 0,
            'cache_misses': 0
        }
        
        self.logger.info("UI Element Detector initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for element detector"""
        logger = logging.getLogger("ElementDetector")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler if debug mode is enabled
            if self.config.get('debugging.debug_mode', False):
                try:
                    log_file = "EHC_Logs/element_detector.log"
                    os.makedirs(os.path.dirname(log_file), exist_ok=True)
                    file_handler = logging.FileHandler(log_file, encoding='utf-8')
                    file_handler.setFormatter(formatter)
                    logger.addHandler(file_handler)
                except Exception:
                    pass
        
        return logger
    
    def detect_element(self, screenshot: np.ndarray, element: ElementDescriptor) -> DetectionResult:
        """Detect a single UI element using multiple methods"""
        start_time = time.time()
        self.stats['total_detections'] += 1
        
        try:
            # Check cache first
            cache_key = self._generate_cache_key(element)
            cached_match = self._get_cached_element(cache_key)
            
            if cached_match:
                self.stats['cache_hits'] += 1
                return DetectionResult(
                    success=True,
                    matches=[cached_match],
                    total_time=time.time() - start_time,
                    methods_tried=["cache"]
                )
            
            self.stats['cache_misses'] += 1
            
            # Extract region if specified
            roi = screenshot
            region_offset = (0, 0)
            
            if element.region and self.region_optimization:
                x, y, w, h = element.region
                roi = screenshot[y:y+h, x:x+w]
                region_offset = (x, y)
            
            # Determine detection methods to try
            methods_to_try = self._get_detection_methods(element)
            methods_tried = []
            all_matches = []
            
            for method in methods_to_try:
                method_start = time.time()
                methods_tried.append(method.value)
                
                try:
                    if method == DetectionMethod.OCR:
                        matches = self._detect_by_ocr(roi, element, region_offset)
                    elif method == DetectionMethod.TEMPLATE:
                        matches = self._detect_by_template(roi, element, region_offset)
                    elif method == DetectionMethod.HYBRID:
                        matches = self._detect_by_hybrid(roi, element, region_offset)
                    else:
                        continue
                    
                    detection_time = time.time() - method_start
                    
                    # Add detection time to matches
                    for match in matches:
                        match.detection_time = detection_time
                    
                    all_matches.extend(matches)
                    self.stats['method_usage'][method.value] += 1
                    
                    self.logger.debug(f"Method {method.value} found {len(matches)} matches for {element.name}")
                    
                except Exception as e:
                    self.logger.warning(f"Method {method.value} failed for {element.name}: {e}")
                    continue
            
            # Filter and rank matches
            valid_matches = [m for m in all_matches if m.confidence >= element.confidence_threshold]
            valid_matches.sort(key=lambda m: m.confidence, reverse=True)
            
            total_time = time.time() - start_time
            
            if valid_matches:
                self.stats['successful_detections'] += 1
                
                # Cache the best match
                if self.cache_enabled and cache_key:
                    self._cache_element(cache_key, valid_matches[0])
                
                # Update average detection time
                self._update_average_time(total_time)
                
                return DetectionResult(
                    success=True,
                    matches=valid_matches,
                    total_time=total_time,
                    methods_tried=methods_tried
                )
            else:
                return DetectionResult(
                    success=False,
                    matches=[],
                    total_time=total_time,
                    methods_tried=methods_tried,
                    error_message=f"Element '{element.name}' not found with sufficient confidence"
                )
                
        except Exception as e:
            total_time = time.time() - start_time
            self.logger.error(f"Element detection failed for {element.name}: {e}")
            
            return DetectionResult(
                success=False,
                matches=[],
                total_time=total_time,
                methods_tried=methods_tried if 'methods_tried' in locals() else [],
                error_message=f"Detection error: {str(e)}"
            )
    
    def detect_multiple_elements(self, screenshot: np.ndarray, elements: List[ElementDescriptor]) -> Dict[str, DetectionResult]:
        """Detect multiple UI elements efficiently"""
        results = {}
        
        for element in elements:
            result = self.detect_element(screenshot, element)
            results[element.name] = result
            
            if result.success:
                self.logger.info(f"Found element '{element.name}' using {result.methods_tried}")
            else:
                self.logger.warning(f"Failed to find element '{element.name}': {result.error_message}")
        
        return results
    
    def _detect_by_ocr(self, screenshot: np.ndarray, element: ElementDescriptor, offset: Tuple[int, int]) -> List[ElementMatch]:
        """Detect element using OCR text recognition"""
        matches = []
        
        if not element.text_patterns:
            return matches
        
        try:
            for text_pattern in element.text_patterns:
                text_matches = self.ocr_service.find_text(screenshot, text_pattern, case_sensitive=False)
                
                for text_match in text_matches:
                    # Adjust coordinates for region offset
                    adjusted_center = (text_match.center[0] + offset[0], text_match.center[1] + offset[1])
                    adjusted_bbox = (
                        text_match.bbox[0] + offset[0],
                        text_match.bbox[1] + offset[1],
                        text_match.bbox[2],
                        text_match.bbox[3]
                    )
                    
                    # Calculate clickable region with offset
                    clickable_center = (
                        adjusted_center[0] + element.clickable_offset[0],
                        adjusted_center[1] + element.clickable_offset[1]
                    )
                    
                    clickable_region = (
                        adjusted_bbox[0] - 5,
                        adjusted_bbox[1] - 5,
                        adjusted_bbox[2] + 10,
                        adjusted_bbox[3] + 10
                    )
                    
                    match = ElementMatch(
                        element_name=element.name,
                        method_used="ocr",
                        confidence=text_match.confidence,
                        location=clickable_center,
                        bounding_box=adjusted_bbox,
                        clickable_region=clickable_region,
                        detection_time=0,  # Will be set by caller
                        raw_match=text_match
                    )
                    matches.append(match)
            
        except Exception as e:
            self.logger.error(f"OCR detection failed for {element.name}: {e}")
        
        return matches
    
    def _detect_by_template(self, screenshot: np.ndarray, element: ElementDescriptor, offset: Tuple[int, int]) -> List[ElementMatch]:
        """Detect element using template matching"""
        matches = []
        
        if not element.template_names:
            return matches
        
        try:
            for template_name in element.template_names:
                template_result = self.template_service.find_template(screenshot, template_name)
                
                if template_result.success:
                    for template_match in template_result.matches:
                        # Calculate center and adjust for offset
                        center = (
                            template_match.location[0] + template_match.size[0] // 2 + offset[0],
                            template_match.location[1] + template_match.size[1] // 2 + offset[1]
                        )
                        
                        # Adjust bounding box for offset
                        bbox = (
                            template_match.location[0] + offset[0],
                            template_match.location[1] + offset[1],
                            template_match.size[0],
                            template_match.size[1]
                        )
                        
                        # Calculate clickable region with offset
                        clickable_center = (
                            center[0] + element.clickable_offset[0],
                            center[1] + element.clickable_offset[1]
                        )
                        
                        clickable_region = (
                            bbox[0] - 5,
                            bbox[1] - 5,
                            bbox[2] + 10,
                            bbox[3] + 10
                        )
                        
                        match = ElementMatch(
                            element_name=element.name,
                            method_used="template",
                            confidence=template_match.confidence,
                            location=clickable_center,
                            bounding_box=bbox,
                            clickable_region=clickable_region,
                            detection_time=0,  # Will be set by caller
                            raw_match=template_match
                        )
                        matches.append(match)
            
        except Exception as e:
            self.logger.error(f"Template detection failed for {element.name}: {e}")
        
        return matches
    
    def _detect_by_hybrid(self, screenshot: np.ndarray, element: ElementDescriptor, offset: Tuple[int, int]) -> List[ElementMatch]:
        """Detect element using hybrid OCR + template approach"""
        matches = []
        
        try:
            # Get matches from both methods
            ocr_matches = self._detect_by_ocr(screenshot, element, offset)
            template_matches = self._detect_by_template(screenshot, element, offset)
            
            # Combine and validate matches
            all_matches = ocr_matches + template_matches
            
            # Remove duplicates based on proximity
            filtered_matches = self._remove_duplicate_matches(all_matches)
            
            # Boost confidence for matches found by both methods
            for match in filtered_matches:
                match.method_used = "hybrid"
                
                # Check if there's a corresponding match from the other method
                other_matches = template_matches if match.method_used == "ocr" else ocr_matches
                for other_match in other_matches:
                    distance = self._calculate_distance(match.location, other_match.location)
                    if distance < 20:  # Within 20 pixels
                        # Boost confidence for correlated matches
                        match.confidence = min(1.0, match.confidence * 1.2)
                        break
            
            matches = filtered_matches
            
        except Exception as e:
            self.logger.error(f"Hybrid detection failed for {element.name}: {e}")
        
        return matches
    
    def _get_detection_methods(self, element: ElementDescriptor) -> List[DetectionMethod]:
        """Determine which detection methods to use for an element"""
        methods = []
        
        if element.method_preference:
            methods.append(element.method_preference)
        else:
            # Default priority: OCR -> Template -> Hybrid
            if element.text_patterns:
                methods.append(DetectionMethod.OCR)
            
            if element.template_names:
                methods.append(DetectionMethod.TEMPLATE)
            
            # Add hybrid if both text and template are available
            if element.text_patterns and element.template_names:
                methods.append(DetectionMethod.HYBRID)
        
        return methods
    
    def _remove_duplicate_matches(self, matches: List[ElementMatch], proximity_threshold: int = 20) -> List[ElementMatch]:
        """Remove duplicate matches that are too close to each other"""
        if not matches:
            return matches
        
        # Sort by confidence (highest first)
        sorted_matches = sorted(matches, key=lambda m: m.confidence, reverse=True)
        filtered_matches = []
        
        for match in sorted_matches:
            is_duplicate = False
            
            for existing in filtered_matches:
                distance = self._calculate_distance(match.location, existing.location)
                if distance < proximity_threshold:
                    is_duplicate = True
                    break
            
            if not is_duplicate:
                filtered_matches.append(match)
        
        return filtered_matches
    
    def _calculate_distance(self, point1: Tuple[int, int], point2: Tuple[int, int]) -> float:
        """Calculate Euclidean distance between two points"""
        dx = point1[0] - point2[0]
        dy = point1[1] - point2[1]
        return (dx ** 2 + dy ** 2) ** 0.5
    
    def _generate_cache_key(self, element: ElementDescriptor) -> Optional[str]:
        """Generate cache key for element"""
        try:
            key_parts = [element.name]
            
            if element.text_patterns:
                key_parts.extend(sorted(element.text_patterns))
            
            if element.template_names:
                key_parts.extend(sorted(element.template_names))
            
            if element.region:
                key_parts.append(f"region_{element.region}")
            
            return "_".join(key_parts)
            
        except Exception:
            return None
    
    def _get_cached_element(self, cache_key: str) -> Optional[ElementMatch]:
        """Get cached element match if still valid"""
        if not cache_key or cache_key not in self.element_cache:
            return None
        
        cached_data = self.element_cache[cache_key]
        if time.time() - cached_data['timestamp'] > self.cache_duration:
            # Cache expired
            del self.element_cache[cache_key]
            return None
        
        return cached_data['match']
    
    def _cache_element(self, cache_key: str, match: ElementMatch):
        """Cache successful element match"""
        if cache_key:
            self.element_cache[cache_key] = {
                'match': match,
                'timestamp': time.time()
            }
    
    def _update_average_time(self, detection_time: float):
        """Update average detection time"""
        if self.stats['total_detections'] > 0:
            current_avg = self.stats['average_detection_time']
            self.stats['average_detection_time'] = (
                (current_avg * (self.stats['total_detections'] - 1) + detection_time) / 
                self.stats['total_detections']
            )
    
    def create_vbs_element_descriptors(self) -> Dict[str, ElementDescriptor]:
        """Create predefined element descriptors for VBS UI elements"""
        return {
            "sales_distribution_menu": ElementDescriptor(
                name="sales_distribution_menu",
                text_patterns=["Sales & Distribution", "Sales and Distribution", "Sales"],
                template_names=["sales_distribution_menu", "sales_menu"],
                confidence_threshold=0.7
            ),
            "pos_menu": ElementDescriptor(
                name="pos_menu",
                text_patterns=["POS", "Point of Sale"],
                template_names=["pos_menu", "pos_button"],
                confidence_threshold=0.7
            ),
            "wifi_registration_menu": ElementDescriptor(
                name="wifi_registration_menu",
                text_patterns=["WiFi User Registration", "WiFi Registration", "User Registration"],
                template_names=["wifi_registration_menu", "wifi_menu"],
                confidence_threshold=0.7
            ),
            "import_button": ElementDescriptor(
                name="import_button",
                text_patterns=["Import", "IMPORT"],
                template_names=["import_button", "import_btn"],
                confidence_threshold=0.8
            ),
            "browse_button": ElementDescriptor(
                name="browse_button",
                text_patterns=["Browse", "Browse..."],
                template_names=["browse_button", "browse_btn"],
                confidence_threshold=0.8
            ),
            "update_button": ElementDescriptor(
                name="update_button",
                text_patterns=["Update", "UPDATE"],
                template_names=["update_button", "update_btn"],
                confidence_threshold=0.8
            ),
            "ok_button": ElementDescriptor(
                name="ok_button",
                text_patterns=["OK", "Ok"],
                template_names=["ok_button", "ok_btn"],
                confidence_threshold=0.8
            ),
            "cancel_button": ElementDescriptor(
                name="cancel_button",
                text_patterns=["Cancel", "CANCEL"],
                template_names=["cancel_button", "cancel_btn"],
                confidence_threshold=0.8
            ),
            "reports_menu": ElementDescriptor(
                name="reports_menu",
                text_patterns=["Reports", "REPORTS"],
                template_names=["reports_menu", "reports_button"],
                confidence_threshold=0.7
            ),
            "print_button": ElementDescriptor(
                name="print_button",
                text_patterns=["Print", "PRINT"],
                template_names=["print_button", "print_btn"],
                confidence_threshold=0.8
            ),
            "export_button": ElementDescriptor(
                name="export_button",
                text_patterns=["Export", "EXPORT"],
                template_names=["export_button", "export_btn"],
                confidence_threshold=0.8
            )
        }
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get element detection performance statistics"""
        stats = self.stats.copy()
        
        # Calculate success rate
        if stats['total_detections'] > 0:
            stats['success_rate'] = stats['successful_detections'] / stats['total_detections']
        else:
            stats['success_rate'] = 0.0
        
        # Add cache efficiency
        total_cache_operations = stats['cache_hits'] + stats['cache_misses']
        if total_cache_operations > 0:
            stats['cache_hit_rate'] = stats['cache_hits'] / total_cache_operations
        else:
            stats['cache_hit_rate'] = 0.0
        
        return stats
    
    def reset_stats(self):
        """Reset performance statistics"""
        self.stats = {
            'total_detections': 0,
            'successful_detections': 0,
            'method_usage': {method.value: 0 for method in DetectionMethod},
            'average_detection_time': 0.0,
            'cache_hits': 0,
            'cache_misses': 0
        }
        
        # Clear cache
        self.element_cache.clear()
    
    def clear_cache(self):
        """Clear element cache"""
        self.element_cache.clear()
        self.logger.info("Element cache cleared")

if __name__ == "__main__":
    # Test element detector
    detector = UIElementDetector()
    print("UI Element Detector initialized successfully")
    
    # Get VBS element descriptors
    vbs_elements = detector.create_vbs_element_descriptors()
    print(f"Created {len(vbs_elements)} VBS element descriptors")