#!/usr/bin/env python3
"""
Test script for Template Service
"""

import sys
import os
import time
import cv2
import numpy as np

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from template_service import TemplateService

def test_template_service():
    """Test template service functionality"""
    print("Testing Template Service...")
    
    try:
        # Initialize service
        service = TemplateService()
        print(f"✓ Template service initialized")
        
        # Test template loading
        templates = service.get_template_list()
        print(f"✓ Found {len(templates)} templates: {templates}")
        
        # Test screenshot capture
        screenshot = service.capture_screenshot()
        if screenshot is not None:
            print(f"✓ Screenshot captured: {screenshot.shape}")
        else:
            print("⚠ Screenshot capture failed")
        
        # Test image preprocessing
        if screenshot is not None:
            processed = service.preprocess_image(screenshot)
            print(f"✓ Image preprocessing: {processed.shape}")
        
        # Test template matching (if templates exist)
        if templates:
            template_name = templates[0]
            print(f"\nTesting template matching with: {template_name}")
            
            result = service.match_template(template_name)
            print(f"✓ Template matching completed")
            print(f"  Success: {result.success}")
            print(f"  Matches: {len(result.matches)}")
            print(f"  Processing time: {result.processing_time:.3f}s")
            
            if result.matches:
                match = result.matches[0]
                print(f"  Best match confidence: {match.confidence:.3f}")
                print(f"  Location: {match.location}")
                print(f"  Center: {match.center}")
        
        # Test performance stats
        stats = service.get_performance_stats()
        print(f"\n✓ Performance stats:")
        print(f"  Total matches: {stats['total_matches']}")
        print(f"  Successful matches: {stats['successful_matches']}")
        print(f"  Average processing time: {stats['average_processing_time']:.3f}s")
        
        print("\n✓ All tests completed successfully!")
        return True
        
    except Exception as e:
        print(f"✗ Test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_test_template():
    """Create a simple test template"""
    try:
        # Create a simple test template (red rectangle)
        template = np.zeros((50, 100, 3), dtype=np.uint8)
        template[:, :] = (0, 0, 255)  # Red color
        
        # Add some text
        cv2.putText(template, "TEST", (20, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 255, 255), 2)
        
        # Save template
        template_dir = "vbs/templates"
        os.makedirs(template_dir, exist_ok=True)
        
        cv2.imwrite(os.path.join(template_dir, "test_template.png"), template)
        
        # Create metadata
        import json
        metadata = {
            'created_date': time.time(),
            'usage_count': 0,
            'success_rate': 0.0,
            'confidence_threshold': 0.8,
            'scale_factors': [1.0],
            'description': "Test template for validation"
        }
        
        with open(os.path.join(template_dir, "test_template.json"), 'w') as f:
            json.dump(metadata, f, indent=2)
        
        print("✓ Test template created")
        return True
        
    except Exception as e:
        print(f"✗ Failed to create test template: {str(e)}")
        return False

if __name__ == "__main__":
    print("VBS Template Service Test")
    print("=" * 40)
    
    # Create test template if needed
    if not os.path.exists("vbs/templates/test_template.png"):
        create_test_template()
    
    # Run tests
    success = test_template_service()
    
    if success:
        print("\n🎉 Template service is working correctly!")
    else:
        print("\n❌ Template service has issues!")
        sys.exit(1)