# VBS OpenCV Modernization - Troubleshooting Guide

## Quick Diagnostic Checklist

### Before You Start
1. ✅ Check if Tesseract OCR is installed and accessible
2. ✅ Verify OpenCV and required Python packages are installed
3. ✅ Confirm VBS application is running and accessible
4. ✅ Check if template images exist in `vbs/templates/` directory
5. ✅ Validate configuration file `config/cv_config.json`

## Common Issues and Solutions

### 1. OCR Not Detecting Text

#### Symptoms
- OCR service returns empty results
- Text elements not found despite being visible
- Low confidence scores for text detection

#### Diagnostic Steps
```python
# Enable debug mode to see processed images
config = {
    "debugging": {
        "debug_mode": True,
        "save_debug_images": True
    }
}

# Check OCR preprocessing
ocr_service = OCRService()
screenshot = capture_screenshot()
processed = ocr_service.preprocess_image_for_ocr(screenshot)
cv2.imwrite("debug_preprocessed.png", processed)
```

#### Solutions

**A. Adjust OCR Preprocessing Parameters**
```json
{
  "ocr_settings": {
    "preprocessing": {
      "gaussian_blur_kernel": 5,        // Increase for noisy images
      "adaptive_threshold_block_size": 15,  // Larger for bigger text
      "adaptive_threshold_c": 4,        // Adjust contrast
      "morphology_kernel_size": 3       // Clean up text edges
    }
  }
}
```

**B. Lower Confidence Threshold**
```json
{
  "ocr_settings": {
    "confidence_threshold": 0.5  // Lower from default 0.7
  }
}
```

**C. Adjust Character Whitelist**
```json
{
  "ocr_settings": {
    "character_whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .-&"
  }
}
```

**D. Try Different Page Segmentation Modes**
```json
{
  "ocr_settings": {
    "page_segmentation_mode": 8  // Try 6, 7, 8, or 13
  }
}
```

### 2. Template Matching Failures

#### Symptoms
- Template matching returns no results
- Low confidence scores for template detection
- Templates not found despite UI elements being visible

#### Diagnostic Steps
```python
# Test template matching with debug output
template_service = TemplateService()
result = template_service.find_template(screenshot, "import_button")
print(f"Template found: {result.success}")
if result.matches:
    for match in result.matches:
        print(f"Confidence: {match.confidence:.2f} at {match.location}")
```

#### Solutions

**A. Update Template Images**
1. Capture new screenshots of UI elements
2. Save as PNG format in `vbs/templates/` directory
3. Use consistent naming convention

**B. Add Multiple Template Variations**
```
vbs/templates/
├── import_button_normal.png
├── import_button_hover.png
├── import_button_pressed.png
└── import_button_disabled.png
```

**C. Adjust Template Matching Parameters**
```json
{
  "template_matching": {
    "confidence_threshold": 0.6,  // Lower threshold
    "scale_factors": [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3],  // More scales
    "max_template_variations": 10  // Allow more variations
  }
}
```

**D. Check Template Image Quality**
- Ensure templates are clear and unambiguous
- Avoid templates with changing elements (like timestamps)
- Use templates that are unique to specific UI elements

### 3. Performance Issues

#### Symptoms
- Slow automation execution
- High memory usage
- Timeouts during operations

#### Diagnostic Steps
```python
# Monitor performance
stats = smart_engine.get_performance_stats()
print(f"Average OCR time: {stats['average_execution_times']['ocr']:.2f}s")
print(f"Success rate: {stats['overall_success_rate']:.2%}")
print(f"Cache hit rate: {stats['cache_hit_rate']:.2%}")
```

#### Solutions

**A. Enable Region-of-Interest Optimization**
```json
{
  "performance": {
    "screenshot_region_optimization": true,
    "cache_successful_locations": true,
    "cache_duration_seconds": 300
  }
}
```

**B. Optimize OCR Processing**
```python
# Use specific regions for OCR instead of full screen
region = (100, 100, 400, 200)  # x, y, width, height
matches = ocr_service.find_text(screenshot, "target_text", region=region)
```

**C. Reduce Template Variations**
```json
{
  "template_matching": {
    "max_template_variations": 3,  // Reduce from default 5
    "scale_factors": [0.9, 1.0, 1.1]  // Fewer scale factors
  }
}
```

**D. Enable Parallel Processing**
```json
{
  "performance": {
    "parallel_processing": true,
    "max_concurrent_operations": 2
  }
}
```

### 4. Error Recovery Not Working

#### Symptoms
- System doesn't recover from failures automatically
- Repeated errors without adaptation
- No parameter adjustments occurring

#### Diagnostic Steps
```python
# Check error handler statistics
error_stats = error_handler.get_error_statistics()
print(f"Recovery success rate: {error_stats['recovery_success_rate']:.2%}")
print(f"Most common errors: {error_stats['most_common_errors']}")

# Export detailed error analysis
error_handler.export_error_analysis("error_analysis.json")
```

#### Solutions

**A. Enable Error Recovery Features**
```json
{
  "smart_automation": {
    "screenshot_on_error": true,
    "performance_tracking": true,
    "exponential_backoff": true,
    "max_retries": 5
  }
}
```

**B. Check Recovery Strategy Configuration**
```python
# Verify recovery strategies are properly initialized
recovery_strategies = error_handler._initialize_recovery_strategies()
print(f"Available strategies: {list(recovery_strategies.keys())}")
```

**C. Enable Automatic Parameter Adjustment**
```json
{
  "debugging": {
    "auto_parameter_adjustment": true,
    "parameter_adjustment_history": true
  }
}
```

### 5. Window Focus Issues

#### Symptoms
- Automation fails to interact with VBS window
- Screenshots capture wrong window
- Click events don't register

#### Diagnostic Steps
```python
# Check window detection
window_handle = smart_engine._find_main_vbs_window()
if window_handle:
    window_title = win32gui.GetWindowText(window_handle)
    print(f"Found VBS window: {window_title}")
else:
    print("No VBS window found")
```

#### Solutions

**A. Improve Window Detection**
```json
{
  "vbs_specific": {
    "window_title_patterns": ["absons", "arabian", "moonflower", "erp", "vbs"],
    "exclude_window_patterns": ["login", "security", "warning", "browser"]
  }
}
```

**B. Force Window Focus**
```python
# Manually set window focus
smart_engine.set_target_window("absons")  # Use specific window title pattern
```

**C. Add Window Validation**
```python
# Validate window before operations
if not smart_engine._ensure_window_focus():
    print("Failed to focus VBS window")
    # Handle error or retry
```

### 6. Configuration Issues

#### Symptoms
- Services fail to initialize
- Invalid configuration errors
- Missing configuration parameters

#### Diagnostic Steps
```python
# Validate configuration
from cv_services.config_loader import get_cv_config
config = get_cv_config()
print(f"Configuration loaded: {config is not None}")

# Check specific settings
ocr_config = config.get_ocr_config()
print(f"OCR settings: {ocr_config}")
```

#### Solutions

**A. Validate Configuration File**
```bash
# Check JSON syntax
python -m json.tool config/cv_config.json
```

**B. Reset to Default Configuration**
```python
# Create default configuration
default_config = {
    "ocr_settings": {
        "confidence_threshold": 0.7,
        "language": "eng",
        "page_segmentation_mode": 6
    },
    "template_matching": {
        "confidence_threshold": 0.8,
        "template_directory": "vbs/templates"
    }
}
```

**C. Check File Permissions**
- Ensure configuration file is readable
- Verify template directory exists and is accessible
- Check log directory permissions

### 7. Installation and Dependencies

#### Symptoms
- Import errors for CV services
- Tesseract not found errors
- OpenCV installation issues

#### Solutions

**A. Install Required Packages**
```bash
pip install opencv-python==4.8.1.78
pip install pytesseract==0.3.10
pip install pillow==10.0.1
pip install numpy
```

**B. Install Tesseract OCR**
1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install to default location: `C:\Program Files\Tesseract-OCR\`
3. Add to system PATH or configure in cv_config.json

**C. Verify Installation**
```python
# Test imports
try:
    import cv2
    import pytesseract
    from cv_services.ocr_service import OCRService
    print("All dependencies installed successfully")
except ImportError as e:
    print(f"Missing dependency: {e}")
```

## Advanced Troubleshooting

### Debug Mode Setup

Enable comprehensive debugging:

```json
{
  "debugging": {
    "debug_mode": true,
    "save_debug_images": true,
    "debug_image_path": "debug_images",
    "verbose_logging": true,
    "screenshot_failed_operations": true
  }
}
```

### Log Analysis

Check log files for detailed information:

```bash
# OCR service logs
tail -f EHC_Logs/ocr_service.log

# Smart engine logs
tail -f EHC_Logs/smart_engine.log

# Error handler logs
tail -f EHC_Logs/cv_error_handler.log
```

### Performance Profiling

Monitor system performance:

```python
import psutil
import time

# Monitor memory usage
process = psutil.Process()
memory_before = process.memory_info().rss / 1024 / 1024  # MB

# Run automation
start_time = time.time()
result = smart_engine.execute_action(action)
execution_time = time.time() - start_time

memory_after = process.memory_info().rss / 1024 / 1024  # MB
memory_used = memory_after - memory_before

print(f"Execution time: {execution_time:.2f}s")
print(f"Memory used: {memory_used:.2f}MB")
```

### Template Debugging

Visualize template matching results:

```python
import cv2
import numpy as np

def debug_template_matching(screenshot, template_path, threshold=0.8):
    template = cv2.imread(template_path, cv2.IMREAD_COLOR)
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    
    locations = np.where(result >= threshold)
    
    # Draw rectangles around matches
    for pt in zip(*locations[::-1]):
        cv2.rectangle(screenshot, pt, (pt[0] + template.shape[1], pt[1] + template.shape[0]), (0, 255, 0), 2)
    
    cv2.imwrite("debug_template_matches.png", screenshot)
    print(f"Found {len(locations[0])} matches above threshold {threshold}")
```

### OCR Debugging

Analyze OCR preprocessing:

```python
def debug_ocr_preprocessing(image):
    ocr_service = OCRService()
    
    # Original image
    cv2.imwrite("debug_01_original.png", image)
    
    # Grayscale conversion
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    cv2.imwrite("debug_02_grayscale.png", gray)
    
    # Gaussian blur
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    cv2.imwrite("debug_03_blurred.png", blurred)
    
    # Adaptive threshold
    thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    cv2.imwrite("debug_04_threshold.png", thresh)
    
    # Morphological operations
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    cleaned = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    cv2.imwrite("debug_05_cleaned.png", cleaned)
    
    return cleaned
```

## Error Codes and Messages

### Common Error Codes

| Code | Message | Cause | Solution |
|------|---------|-------|----------|
| OCR_001 | "Tesseract not found" | Tesseract not installed or not in PATH | Install Tesseract OCR |
| OCR_002 | "Low confidence text detection" | OCR confidence below threshold | Lower confidence threshold or improve image quality |
| TPL_001 | "Template not found" | Template image missing | Add template image to templates directory |
| TPL_002 | "Template matching failed" | No matches above confidence threshold | Update template or lower threshold |
| ENG_001 | "Smart engine initialization failed" | CV services not available | Check dependencies and imports |
| WIN_001 | "Window not found" | VBS window not detected | Check window title patterns |
| CFG_001 | "Configuration file not found" | cv_config.json missing | Create configuration file |
| CFG_002 | "Invalid configuration" | JSON syntax error | Validate JSON syntax |

### Error Recovery Actions

The system automatically attempts these recovery actions:

1. **Retry with same method** - Simple retry with delay
2. **Try next method** - Switch from OCR to template matching
3. **Adjust parameters** - Lower confidence thresholds
4. **Wait and retry** - Wait for UI to stabilize
5. **Capture new template** - Update template images
6. **Fallback to coordinates** - Use hardcoded coordinates
7. **Manual intervention** - Request human assistance

## Getting Help

### Diagnostic Information to Collect

When reporting issues, include:

1. **System Information**
   - Windows version
   - Python version
   - Installed package versions

2. **Configuration Files**
   - cv_config.json contents
   - Any custom configuration

3. **Log Files**
   - Recent log entries from EHC_Logs/
   - Error messages and stack traces

4. **Screenshots**
   - VBS application screenshots
   - Debug images if available
   - Error dialogs or messages

5. **Performance Data**
   - Performance statistics from smart engine
   - Memory usage information
   - Execution timing data

### Diagnostic Script

Run this script to collect diagnostic information:

```python
#!/usr/bin/env python3
"""
VBS OpenCV Diagnostic Script
Collects system information for troubleshooting
"""

import sys
import os
import json
import platform
import pkg_resources
from datetime import datetime

def collect_diagnostics():
    diagnostics = {
        "timestamp": datetime.now().isoformat(),
        "system": {
            "platform": platform.platform(),
            "python_version": sys.version,
            "architecture": platform.architecture()
        },
        "packages": {},
        "configuration": {},
        "files": {}
    }
    
    # Check installed packages
    required_packages = ["opencv-python", "pytesseract", "pillow", "numpy"]
    for package in required_packages:
        try:
            version = pkg_resources.get_distribution(package).version
            diagnostics["packages"][package] = version
        except pkg_resources.DistributionNotFound:
            diagnostics["packages"][package] = "NOT INSTALLED"
    
    # Check configuration
    try:
        with open("config/cv_config.json", "r") as f:
            diagnostics["configuration"] = json.load(f)
    except FileNotFoundError:
        diagnostics["configuration"] = "FILE NOT FOUND"
    except json.JSONDecodeError as e:
        diagnostics["configuration"] = f"JSON ERROR: {e}"
    
    # Check important files
    important_files = [
        "vbs/cv_services/ocr_service.py",
        "vbs/cv_services/template_service.py",
        "vbs/cv_services/smart_engine.py",
        "vbs/templates/"
    ]
    
    for file_path in important_files:
        diagnostics["files"][file_path] = os.path.exists(file_path)
    
    # Save diagnostics
    with open("diagnostics.json", "w") as f:
        json.dump(diagnostics, f, indent=2)
    
    print("Diagnostics saved to diagnostics.json")
    return diagnostics

if __name__ == "__main__":
    collect_diagnostics()
```

### Support Contacts

For additional support:

1. Check the main documentation: `docs/VBS_OpenCV_Modernization_Guide.md`
2. Review error logs in `EHC_Logs/` directory
3. Run the diagnostic script above
4. Create a detailed issue report with collected information

## Preventive Maintenance

### Regular Maintenance Tasks

1. **Weekly**
   - Review error logs for patterns
   - Check performance statistics
   - Validate template images

2. **Monthly**
   - Update template images if UI changes
   - Review and optimize configuration
   - Clean up debug images and old logs

3. **Quarterly**
   - Performance benchmarking
   - Dependency updates
   - Configuration backup

### Monitoring Setup

Set up automated monitoring:

```python
# Monitor error rates
def check_error_rates():
    stats = error_handler.get_error_statistics()
    if stats['recovery_success_rate'] < 0.8:
        send_alert("High error rate detected")

# Monitor performance
def check_performance():
    stats = smart_engine.get_performance_stats()
    if stats['average_execution_times']['ocr'] > 5.0:
        send_alert("Slow OCR performance detected")

# Schedule monitoring
import schedule
schedule.every(1).hours.do(check_error_rates)
schedule.every(1).hours.do(check_performance)
```

This troubleshooting guide should help resolve most common issues with the VBS OpenCV modernization system. For complex issues, enable debug mode and collect comprehensive diagnostic information before seeking additional support.