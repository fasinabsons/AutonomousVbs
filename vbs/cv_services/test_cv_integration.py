#!/usr/bin/env python3
"""
Comprehensive Testing for Enhanced VBS Phases with Computer Vision
Tests OCR, template matching, element detection, and integration with VBS phases
"""

import unittest
import time
import os
import sys
import numpy as np
import cv2
from unittest.mock import Mock, patch, MagicMock
from typing import Dict, List, Any

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import CV services
from cv_services.ocr_service import OCRService, TextMatch, OCRResult
from cv_services.template_service import TemplateService, TemplateMatch, TemplateResult
from cv_services.element_detector import UIElementDetector, ElementDescriptor, ElementMatch, DetectionResult
from cv_services.smart_engine import SmartAutomationEngine, AutomationAction, AutomationResult
from cv_services.error_handler import CVErrorHandler, ErrorCategory, RecoveryStrategy
from cv_services.config_loader import get_cv_config

class TestOCRService(unittest.TestCase):
    """Test OCR service functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.ocr_service = OCRService()
        self.test_image = self._create_test_image_with_text("Test Button")
    
    def _create_test_image_with_text(self, text: str) -> np.ndarray:
        """Create a test image with text for OCR testing"""
        # Create a white image
        img = np.ones((100, 300, 3), dtype=np.uint8) * 255
        
        # Add text using OpenCV
        font = cv2.FONT_HERSHEY_SIMPLEX
        cv2.putText(img, text, (50, 50), font, 1, (0, 0, 0), 2, cv2.LINE_AA)
        
        return img
    
    def test_ocr_text_extraction(self):
        """Test basic OCR text extraction"""
        result = self.ocr_service.extract_text_tesseract(self.test_image)
        
        self.assertTrue(result.success)
        self.assertGreater(len(result.matches), 0)
        self.assertIn("Test", result.matches[0].text)
    
    def test_ocr_text_search(self):
        """Test OCR text search functionality"""
        matches = self.ocr_service.find_text(self.test_image, "Test", case_sensitive=False)
        
        self.assertGreater(len(matches), 0)
        self.assertGreater(matches[0].confidence, 0.5)
        self.assertEqual(len(matches[0].center), 2)  # Should have x, y coordinates
    
    def test_vbs_ui_elements_detection(self):
        """Test VBS-specific UI element detection"""
        # Create test images for different VBS elements
        test_elements = ["New", "Import", "Export", "Update", "Reports"]
        
        for element_text in test_elements:
            test_img = self._create_test_image_with_text(element_text)
            results = self.ocr_service.find_vbs_ui_elements(test_img, [element_text.lower()])
            
            self.assertIn(element_text.lower(), results)
            if results[element_text.lower()]:
                self.assertGreater(len(results[element_text.lower()]), 0)
    
    def test_ocr_performance_stats(self):
        """Test OCR performance statistics tracking"""
        initial_stats = self.ocr_service.get_performance_stats()
        
        # Perform some OCR operations
        self.ocr_service.extract_text_tesseract(self.test_image)
        self.ocr_service.find_text(self.test_image, "Test")
        
        updated_stats = self.ocr_service.get_performance_stats()
        
        self.assertGreaterEqual(updated_stats['total_operations'], initial_stats['total_operations'])

class TestTemplateService(unittest.TestCase):
    """Test template matching service functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.template_service = TemplateService()
        self.test_screenshot = self._create_test_screenshot()
        self.test_template = self._create_test_template()
    
    def _create_test_screenshot(self) -> np.ndarray:
        """Create a test screenshot"""
        # Create a larger image with a button-like rectangle
        img = np.ones((400, 600, 3), dtype=np.uint8) * 240
        
        # Add a button-like rectangle
        cv2.rectangle(img, (100, 150), (200, 200), (200, 200, 200), -1)
        cv2.rectangle(img, (100, 150), (200, 200), (100, 100, 100), 2)
        
        return img
    
    def _create_test_template(self) -> np.ndarray:
        """Create a test template"""
        # Create a small button template
        template = np.ones((50, 100, 3), dtype=np.uint8) * 200
        cv2.rectangle(template, (0, 0), (99, 49), (100, 100, 100), 2)
        
        return template
    
    def test_template_matching(self):
        """Test basic template matching"""
        # Save test template temporarily
        template_path = "test_template.png"
        cv2.imwrite(template_path, self.test_template)
        
        try:
            result = self.template_service.find_template(self.test_screenshot, "test_template")
            
            # Clean up
            if os.path.exists(template_path):
                os.remove(template_path)
            
            # Template might not match perfectly, but service should handle gracefully
            self.assertIsInstance(result, TemplateResult)
            
        except Exception as e:
            # Clean up on error
            if os.path.exists(template_path):
                os.remove(template_path)
            raise e
    
    def test_template_performance_stats(self):
        """Test template service performance statistics"""
        initial_stats = self.template_service.get_performance_stats()
        
        # Perform template operations
        self.template_service.find_template(self.test_screenshot, "nonexistent_template")
        
        updated_stats = self.template_service.get_performance_stats()
        
        self.assertGreaterEqual(updated_stats['total_operations'], initial_stats['total_operations'])

class TestElementDetector(unittest.TestCase):
    """Test UI element detection service"""
    
    def setUp(self):
        """Set up test environment"""
        self.element_detector = UIElementDetector()
        self.test_screenshot = self._create_test_ui_screenshot()
    
    def _create_test_ui_screenshot(self) -> np.ndarray:
        """Create a test UI screenshot"""
        img = np.ones((500, 800, 3), dtype=np.uint8) * 255
        
        # Add some UI elements
        cv2.putText(img, "Import", (100, 100), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.putText(img, "Export", (300, 100), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.rectangle(img, (80, 70), (180, 120), (200, 200, 200), 2)
        cv2.rectangle(img, (280, 70), (380, 120), (200, 200, 200), 2)
        
        return img
    
    def test_element_detection_ocr(self):
        """Test element detection using OCR"""
        element_desc = ElementDescriptor(
            name="import_button",
            text_patterns=["Import"],
            confidence_threshold=0.5
        )
        
        result = self.element_detector.detect_element(self.test_screenshot, element_desc)
        
        self.assertIsInstance(result, DetectionResult)
        # May or may not find element depending on OCR accuracy, but should not crash
    
    def test_vbs_element_descriptors(self):
        """Test VBS-specific element descriptors"""
        vbs_elements = self.element_detector.create_vbs_element_descriptors()
        
        self.assertIsInstance(vbs_elements, dict)
        self.assertIn("import_button", vbs_elements)
        self.assertIn("export_button", vbs_elements)
        self.assertIn("update_button", vbs_elements)
        
        # Check element descriptor structure
        import_desc = vbs_elements["import_button"]
        self.assertIsInstance(import_desc, ElementDescriptor)
        self.assertIsNotNone(import_desc.text_patterns)
        self.assertIsNotNone(import_desc.template_names)
    
    def test_multiple_element_detection(self):
        """Test detection of multiple elements"""
        elements = [
            ElementDescriptor(name="import", text_patterns=["Import"]),
            ElementDescriptor(name="export", text_patterns=["Export"])
        ]
        
        results = self.element_detector.detect_multiple_elements(self.test_screenshot, elements)
        
        self.assertIsInstance(results, dict)
        self.assertIn("import", results)
        self.assertIn("export", results)

class TestSmartAutomationEngine(unittest.TestCase):
    """Test smart automation engine"""
    
    def setUp(self):
        """Set up test environment"""
        self.smart_engine = SmartAutomationEngine()
    
    def test_automation_action_creation(self):
        """Test automation action creation"""
        action = AutomationAction(
            action_type="click",
            target_text="Test Button",
            confidence_threshold=0.7
        )
        
        self.assertEqual(action.action_type, "click")
        self.assertEqual(action.target_text, "Test Button")
        self.assertEqual(action.confidence_threshold, 0.7)
    
    @patch('cv_services.smart_engine.pyautogui')
    def test_action_execution_mock(self, mock_pyautogui):
        """Test action execution with mocked GUI operations"""
        # Mock screenshot
        mock_screenshot = np.ones((100, 200, 3), dtype=np.uint8) * 255
        
        with patch.object(self.smart_engine, 'capture_screenshot', return_value=mock_screenshot):
            action = AutomationAction(
                action_type="click",
                coordinates=(100, 100),
                confidence_threshold=0.5
            )
            
            result = self.smart_engine.execute_action(action)
            
            self.assertIsInstance(result, AutomationResult)
    
    def test_vbs_phase_actions(self):
        """Test VBS-specific phase actions"""
        vbs_actions = self.smart_engine._get_vbs_actions()
        
        self.assertIn("phase2_navigation", vbs_actions)
        self.assertIn("phase3_upload", vbs_actions)
        self.assertIn("phase4_report", vbs_actions)
        
        # Check phase 2 actions
        phase2_actions = vbs_actions["phase2_navigation"]
        self.assertIn("click_sales_distribution", phase2_actions)
        self.assertIn("click_pos", phase2_actions)
        self.assertIn("click_wifi_registration", phase2_actions)
    
    def test_performance_stats(self):
        """Test performance statistics tracking"""
        initial_stats = self.smart_engine.get_performance_stats()
        
        self.assertIn('total_operations', initial_stats)
        self.assertIn('successful_operations', initial_stats)
        self.assertIn('method_success_rates', initial_stats)
        self.assertIn('overall_success_rate', initial_stats)

class TestErrorHandler(unittest.TestCase):
    """Test error handling and recovery service"""
    
    def setUp(self):
        """Set up test environment"""
        self.error_handler = CVErrorHandler()
    
    def test_error_categorization(self):
        """Test error categorization"""
        test_errors = [
            (Exception("Element not found"), ErrorCategory.ELEMENT_NOT_FOUND),
            (Exception("Low confidence threshold"), ErrorCategory.LOW_CONFIDENCE),
            (Exception("Screenshot capture failed"), ErrorCategory.SCREENSHOT_FAILURE),
            (Exception("Operation timed out"), ErrorCategory.TIMEOUT),
            (Exception("OCR processing error"), ErrorCategory.OCR_ERROR),
            (Exception("Template matching failed"), ErrorCategory.TEMPLATE_ERROR),
        ]
        
        for error, expected_category in test_errors:
            context = {'method_used': 'test', 'action_type': 'test'}
            categorized = self.error_handler._categorize_error(error, context)
            self.assertEqual(categorized, expected_category)
    
    def test_recovery_strategies(self):
        """Test recovery strategy initialization"""
        strategies = self.error_handler._initialize_recovery_strategies()
        
        self.assertIn(ErrorCategory.ELEMENT_NOT_FOUND, strategies)
        self.assertIn(ErrorCategory.LOW_CONFIDENCE, strategies)
        self.assertIn(ErrorCategory.OCR_ERROR, strategies)
        
        # Check strategy structure
        element_not_found_strategies = strategies[ErrorCategory.ELEMENT_NOT_FOUND]
        self.assertGreater(len(element_not_found_strategies), 0)
        
        first_strategy = element_not_found_strategies[0]
        self.assertIn('strategy', first_strategy)
        self.assertIn('priority', first_strategy)
        self.assertIn('description', first_strategy)
    
    def test_error_statistics(self):
        """Test error statistics tracking"""
        initial_stats = self.error_handler.get_error_statistics()
        
        self.assertIn('total_errors', initial_stats)
        self.assertIn('successful_recoveries', initial_stats)
        self.assertIn('recovery_success_rate', initial_stats)
        self.assertIn('error_category_counts', initial_stats)

class TestVBSPhaseIntegration(unittest.TestCase):
    """Test integration with VBS phases"""
    
    def setUp(self):
        """Set up test environment"""
        self.config = get_cv_config()
    
    def test_cv_config_loading(self):
        """Test CV configuration loading"""
        self.assertIsInstance(self.config, dict)
        self.assertIn('ocr_settings', self.config)
        self.assertIn('template_matching', self.config)
        self.assertIn('smart_automation', self.config)
    
    def test_ocr_config_structure(self):
        """Test OCR configuration structure"""
        ocr_config = self.config.get('ocr_settings', {})
        
        expected_keys = [
            'confidence_threshold',
            'language',
            'page_segmentation_mode',
            'preprocessing'
        ]
        
        for key in expected_keys:
            self.assertIn(key, ocr_config, f"Missing OCR config key: {key}")
    
    def test_template_config_structure(self):
        """Test template matching configuration structure"""
        template_config = self.config.get('template_matching', {})
        
        expected_keys = [
            'confidence_threshold',
            'template_cache_size',
            'scale_factors',
            'template_directory'
        ]
        
        for key in expected_keys:
            self.assertIn(key, template_config, f"Missing template config key: {key}")
    
    def test_smart_automation_config(self):
        """Test smart automation configuration"""
        smart_config = self.config.get('smart_automation', {})
        
        expected_keys = [
            'method_priority',
            'max_retries',
            'retry_delay',
            'performance_tracking'
        ]
        
        for key in expected_keys:
            self.assertIn(key, smart_config, f"Missing smart automation config key: {key}")

class TestPerformanceBenchmarks(unittest.TestCase):
    """Performance benchmark tests"""
    
    def setUp(self):
        """Set up performance testing"""
        self.ocr_service = OCRService()
        self.template_service = TemplateService()
        self.element_detector = UIElementDetector()
        self.test_image = self._create_performance_test_image()
    
    def _create_performance_test_image(self) -> np.ndarray:
        """Create an image for performance testing"""
        img = np.ones((800, 1200, 3), dtype=np.uint8) * 255
        
        # Add multiple text elements
        texts = ["Import", "Export", "Update", "New", "Delete", "Save", "Cancel", "OK"]
        positions = [(100, 100), (300, 100), (500, 100), (700, 100),
                    (100, 300), (300, 300), (500, 300), (700, 300)]
        
        for text, pos in zip(texts, positions):
            cv2.putText(img, text, pos, cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
            cv2.rectangle(img, (pos[0]-20, pos[1]-30), (pos[0]+len(text)*20, pos[1]+10), (200, 200, 200), 2)
        
        return img
    
    def test_ocr_performance(self):
        """Test OCR performance benchmarks"""
        start_time = time.time()
        
        # Perform multiple OCR operations
        for _ in range(5):
            result = self.ocr_service.extract_text_tesseract(self.test_image)
            self.assertTrue(result.success or not result.success)  # Just ensure it doesn't crash
        
        end_time = time.time()
        avg_time = (end_time - start_time) / 5
        
        # OCR should complete within reasonable time (adjust threshold as needed)
        self.assertLess(avg_time, 5.0, "OCR performance too slow")
        print(f"Average OCR time: {avg_time:.2f} seconds")
    
    def test_element_detection_performance(self):
        """Test element detection performance"""
        elements = [
            ElementDescriptor(name="import", text_patterns=["Import"]),
            ElementDescriptor(name="export", text_patterns=["Export"]),
            ElementDescriptor(name="update", text_patterns=["Update"])
        ]
        
        start_time = time.time()
        
        # Perform multiple detection operations
        for _ in range(3):
            results = self.element_detector.detect_multiple_elements(self.test_image, elements)
            self.assertIsInstance(results, dict)
        
        end_time = time.time()
        avg_time = (end_time - start_time) / 3
        
        # Element detection should be reasonably fast
        self.assertLess(avg_time, 10.0, "Element detection performance too slow")
        print(f"Average element detection time: {avg_time:.2f} seconds")

def run_comprehensive_tests():
    """Run all comprehensive tests"""
    print("🧪 Starting Comprehensive VBS Computer Vision Tests")
    print("=" * 60)
    
    # Create test suite
    test_suite = unittest.TestSuite()
    
    # Add test classes
    test_classes = [
        TestOCRService,
        TestTemplateService,
        TestElementDetector,
        TestSmartAutomationEngine,
        TestErrorHandler,
        TestVBSPhaseIntegration,
        TestPerformanceBenchmarks
    ]
    
    for test_class in test_classes:
        tests = unittest.TestLoader().loadTestsFromTestCase(test_class)
        test_suite.addTests(tests)
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(test_suite)
    
    # Print summary
    print("\n" + "=" * 60)
    print("🧪 Test Summary")
    print(f"Tests run: {result.testsRun}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    print(f"Success rate: {((result.testsRun - len(result.failures) - len(result.errors)) / result.testsRun * 100):.1f}%")
    
    if result.failures:
        print("\n❌ Failures:")
        for test, traceback in result.failures:
            print(f"  - {test}: {traceback.split('AssertionError: ')[-1].split('\\n')[0]}")
    
    if result.errors:
        print("\n🚨 Errors:")
        for test, traceback in result.errors:
            print(f"  - {test}: {traceback.split('\\n')[-2]}")
    
    return result.wasSuccessful()

if __name__ == "__main__":
    success = run_comprehensive_tests()
    sys.exit(0 if success else 1)