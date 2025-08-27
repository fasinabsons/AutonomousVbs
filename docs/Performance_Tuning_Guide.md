# VBS OpenCV Performance Tuning Guide

## Overview

This guide provides comprehensive strategies for optimizing the performance of the VBS OpenCV modernization system. The goal is to achieve fast, reliable automation while maintaining accuracy and system stability.

## Performance Metrics

### Key Performance Indicators (KPIs)

1. **Execution Time**: Time to complete automation actions
2. **Success Rate**: Percentage of successful operations
3. **Memory Usage**: RAM consumption during operations
4. **CPU Utilization**: Processor usage patterns
5. **Cache Hit Rate**: Effectiveness of location caching
6. **Error Recovery Rate**: Success rate of automatic error recovery

### Baseline Performance Targets

| Metric | Target | Acceptable | Poor |
|--------|--------|------------|------|
| OCR Detection Time | < 2 seconds | < 5 seconds | > 5 seconds |
| Template Matching Time | < 1 second | < 3 seconds | > 3 seconds |
| Overall Success Rate | > 95% | > 85% | < 85% |
| Memory Usage | < 500 MB | < 1 GB | > 1 GB |
| Cache Hit Rate | > 70% | > 50% | < 50% |
| Error Recovery Rate | > 90% | > 75% | < 75% |

## OCR Performance Optimization

### 1. Image Preprocessing Optimization

#### Optimal Preprocessing Parameters

```json
{
  "ocr_settings": {
    "preprocessing": {
      "gaussian_blur_kernel": 3,        // Start with 3, increase for noisy images
      "adaptive_threshold_block_size": 11,  // Odd numbers only, 11-15 optimal
      "adaptive_threshold_c": 2,        // 2-4 range, adjust for contrast
      "morphology_kernel_size": 2       // 1-3 range, 2 is usually optimal
    }
  }
}
```

#### Performance Impact Analysis

```python
import time
import cv2

def benchmark_preprocessing(image, iterations=100):
    """Benchmark different preprocessing parameters"""
    
    configs = [
        {"blur": 1, "block": 9, "c": 2},
        {"blur": 3, "block": 11, "c": 2},
        {"blur": 5, "block": 13, "c": 3},
        {"blur": 7, "block": 15, "c": 4}
    ]
    
    results = []
    
    for config in configs:
        start_time = time.time()
        
        for _ in range(iterations):
            # Grayscale
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # Gaussian blur
            if config["blur"] > 0:
                blurred = cv2.GaussianBlur(gray, (config["blur"], config["blur"]), 0)
            else:
                blurred = gray
            
            # Adaptive threshold
            thresh = cv2.adaptiveThreshold(
                blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY, config["block"], config["c"]
            )
            
            # Morphological operations
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
            cleaned = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
        
        avg_time = (time.time() - start_time) / iterations
        results.append((config, avg_time))
    
    return results
```

### 2. Region of Interest (ROI) Optimization

#### Strategic ROI Usage

```python
# Define common UI regions for VBS
VBS_REGIONS = {
    "menu_area": (0, 0, 400, 800),           # Left menu area
    "content_area": (400, 0, 1200, 800),     # Main content area
    "button_area": (0, 700, 1600, 100),      # Bottom button area
    "dialog_area": (300, 200, 1000, 400)     # Dialog box area
}

def optimized_ocr_search(screenshot, target_text, region_hint=None):
    """Use ROI to speed up OCR searches"""
    
    if region_hint and region_hint in VBS_REGIONS:
        x, y, w, h = VBS_REGIONS[region_hint]
        roi = screenshot[y:y+h, x:x+w]
        
        # Search in ROI first
        matches = ocr_service.find_text(roi, target_text)
        
        # Adjust coordinates back to full screen
        for match in matches:
            match.center = (match.center[0] + x, match.center[1] + y)
            match.bbox = (match.bbox[0] + x, match.bbox[1] + y, match.bbox[2], match.bbox[3])
        
        if matches:
            return matches
    
    # Fallback to full screen search
    return ocr_service.find_text(screenshot, target_text)
```

### 3. OCR Engine Optimization

#### Tesseract Configuration Tuning

```json
{
  "ocr_settings": {
    "page_segmentation_mode": 6,    // 6 = uniform block (fastest for UI text)
    "ocr_engine_mode": 3,           // 3 = default (balanced speed/accuracy)
    "character_whitelist": "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .-&",
    "dpi": 96,                      // Match screen DPI
    "timeout": 10                   // Prevent hanging on difficult images
  }
}
```

#### Multi-threading OCR

```python
import concurrent.futures
import threading

class ThreadedOCRService:
    def __init__(self, max_workers=2):
        self.max_workers = max_workers
        self.thread_local = threading.local()
    
    def get_ocr_engine(self):
        """Get thread-local OCR engine"""
        if not hasattr(self.thread_local, 'ocr_service'):
            self.thread_local.ocr_service = OCRService()
        return self.thread_local.ocr_service
    
    def parallel_text_search(self, screenshot, text_patterns):
        """Search for multiple text patterns in parallel"""
        
        def search_text(pattern):
            ocr_service = self.get_ocr_engine()
            return pattern, ocr_service.find_text(screenshot, pattern)
        
        results = {}
        with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_to_pattern = {executor.submit(search_text, pattern): pattern 
                               for pattern in text_patterns}
            
            for future in concurrent.futures.as_completed(future_to_pattern):
                pattern, matches = future.result()
                results[pattern] = matches
        
        return results
```

## Template Matching Optimization

### 1. Template Management

#### Optimal Template Characteristics

```python
def optimize_template(template_image):
    """Optimize template for better matching performance"""
    
    # Resize if too large (max 200x200 for performance)
    height, width = template_image.shape[:2]
    if width > 200 or height > 200:
        scale = min(200/width, 200/height)
        new_width = int(width * scale)
        new_height = int(height * scale)
        template_image = cv2.resize(template_image, (new_width, new_height))
    
    # Convert to grayscale for faster matching
    if len(template_image.shape) == 3:
        template_image = cv2.cvtColor(template_image, cv2.COLOR_BGR2GRAY)
    
    # Enhance contrast
    template_image = cv2.equalizeHist(template_image)
    
    return template_image
```

#### Template Caching Strategy

```python
class OptimizedTemplateService:
    def __init__(self):
        self.template_cache = {}
        self.cache_stats = {"hits": 0, "misses": 0}
        self.max_cache_size = 50
    
    def get_template(self, template_name):
        """Get template with LRU caching"""
        
        if template_name in self.template_cache:
            self.cache_stats["hits"] += 1
            # Move to end (most recently used)
            template = self.template_cache.pop(template_name)
            self.template_cache[template_name] = template
            return template
        
        self.cache_stats["misses"] += 1
        
        # Load and optimize template
        template_path = f"vbs/templates/{template_name}.png"
        template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
        
        if template is not None:
            template = self.optimize_template(template)
            
            # Add to cache (remove oldest if full)
            if len(self.template_cache) >= self.max_cache_size:
                oldest = next(iter(self.template_cache))
                del self.template_cache[oldest]
            
            self.template_cache[template_name] = template
        
        return template
```

### 2. Multi-Scale Matching Optimization

#### Intelligent Scale Selection

```python
def adaptive_scale_matching(screenshot, template, base_scales=[0.8, 0.9, 1.0, 1.1, 1.2]):
    """Adaptive scale matching with early termination"""
    
    best_match = None
    best_confidence = 0
    
    # Try scales in order of likelihood (1.0 first, then nearby scales)
    ordered_scales = [1.0, 0.9, 1.1, 0.8, 1.2]
    
    for scale in ordered_scales:
        if scale not in base_scales:
            continue
            
        # Resize template
        if scale != 1.0:
            height, width = template.shape[:2]
            new_height, new_width = int(height * scale), int(width * scale)
            scaled_template = cv2.resize(template, (new_width, new_height))
        else:
            scaled_template = template
        
        # Perform matching
        result = cv2.matchTemplate(screenshot, scaled_template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)
        
        if max_val > best_confidence:
            best_confidence = max_val
            best_match = {
                "location": max_loc,
                "confidence": max_val,
                "scale": scale,
                "size": (scaled_template.shape[1], scaled_template.shape[0])
            }
            
            # Early termination if confidence is high enough
            if max_val > 0.9:
                break
    
    return best_match
```

### 3. Template Matching Algorithms

#### Algorithm Selection Based on Use Case

```python
def select_matching_method(template_size, screenshot_size, accuracy_requirement):
    """Select optimal template matching method"""
    
    template_pixels = template_size[0] * template_size[1]
    screenshot_pixels = screenshot_size[0] * screenshot_size[1]
    
    # For small templates on large screenshots, use faster methods
    if template_pixels < 10000 and screenshot_pixels > 1000000:
        if accuracy_requirement == "high":
            return cv2.TM_CCOEFF_NORMED  # Most accurate but slower
        else:
            return cv2.TM_CCORR_NORMED   # Faster, good for high contrast
    
    # For similar sized images, use correlation methods
    elif template_pixels > screenshot_pixels * 0.1:
        return cv2.TM_CCOEFF_NORMED
    
    # Default to normalized correlation coefficient
    else:
        return cv2.TM_CCOEFF_NORMED
```

## Memory Management

### 1. Image Memory Optimization

#### Efficient Image Handling

```python
import gc
import weakref

class MemoryEfficientImageProcessor:
    def __init__(self):
        self.image_refs = weakref.WeakSet()
    
    def process_image(self, image_path):
        """Process image with automatic memory cleanup"""
        
        try:
            # Load image
            image = cv2.imread(image_path)
            self.image_refs.add(image)
            
            # Process image
            processed = self.preprocess_image(image)
            
            # Explicitly delete original to free memory
            del image
            
            return processed
            
        finally:
            # Force garbage collection
            gc.collect()
    
    def cleanup_images(self):
        """Force cleanup of all image references"""
        for img_ref in list(self.image_refs):
            try:
                del img_ref
            except:
                pass
        
        self.image_refs.clear()
        gc.collect()
```

### 2. Cache Management

#### Intelligent Cache Sizing

```python
import psutil

class AdaptiveCacheManager:
    def __init__(self):
        self.max_memory_percent = 0.3  # Use max 30% of available RAM
        self.cache = {}
        self.access_times = {}
    
    def get_max_cache_size(self):
        """Calculate maximum cache size based on available memory"""
        available_memory = psutil.virtual_memory().available
        max_cache_memory = available_memory * self.max_memory_percent
        
        # Estimate average item size (adjust based on your data)
        avg_item_size = 1024 * 1024  # 1MB per cached item
        
        return int(max_cache_memory / avg_item_size)
    
    def adaptive_cleanup(self):
        """Clean up cache based on memory pressure"""
        current_memory = psutil.virtual_memory().percent
        
        if current_memory > 80:  # High memory usage
            # Remove 50% of least recently used items
            items_to_remove = len(self.cache) // 2
            sorted_items = sorted(self.access_times.items(), key=lambda x: x[1])
            
            for key, _ in sorted_items[:items_to_remove]:
                if key in self.cache:
                    del self.cache[key]
                    del self.access_times[key]
```

## System-Level Optimizations

### 1. Process Priority and Affinity

#### Windows Process Optimization

```python
import psutil
import os

def optimize_process_priority():
    """Optimize process priority for better performance"""
    
    try:
        # Get current process
        process = psutil.Process(os.getpid())
        
        # Set high priority (be careful with this)
        process.nice(psutil.HIGH_PRIORITY_CLASS)
        
        # Set CPU affinity to use specific cores (optional)
        # Leave at least one core for system processes
        cpu_count = psutil.cpu_count()
        if cpu_count > 2:
            # Use all but the first core
            process.cpu_affinity(list(range(1, cpu_count)))
        
        print(f"Process optimized: Priority={process.nice()}, Affinity={process.cpu_affinity()}")
        
    except Exception as e:
        print(f"Failed to optimize process: {e}")
```

### 2. Threading and Concurrency

#### Optimal Threading Configuration

```python
import threading
from concurrent.futures import ThreadPoolExecutor

class OptimizedAutomationEngine:
    def __init__(self):
        # Determine optimal thread count
        cpu_count = os.cpu_count()
        self.max_workers = min(cpu_count, 4)  # Cap at 4 for UI automation
        
        # Thread pool for parallel operations
        self.executor = ThreadPoolExecutor(max_workers=self.max_workers)
        
        # Thread-local storage for services
        self.thread_local = threading.local()
    
    def parallel_element_detection(self, screenshot, elements):
        """Detect multiple elements in parallel"""
        
        def detect_element(element_desc):
            # Get thread-local detector
            if not hasattr(self.thread_local, 'detector'):
                self.thread_local.detector = UIElementDetector()
            
            return self.thread_local.detector.detect_element(screenshot, element_desc)
        
        # Submit all detection tasks
        futures = {self.executor.submit(detect_element, elem): elem.name 
                  for elem in elements}
        
        results = {}
        for future in futures:
            element_name = futures[future]
            try:
                results[element_name] = future.result(timeout=10)
            except Exception as e:
                results[element_name] = None
                print(f"Element detection failed for {element_name}: {e}")
        
        return results
```

## Configuration Optimization

### 1. Performance-Oriented Configuration

#### High-Performance Configuration

```json
{
  "ocr_settings": {
    "confidence_threshold": 0.6,
    "page_segmentation_mode": 6,
    "ocr_engine_mode": 3,
    "preprocessing": {
      "gaussian_blur_kernel": 3,
      "adaptive_threshold_block_size": 11,
      "adaptive_threshold_c": 2,
      "morphology_kernel_size": 2
    }
  },
  "template_matching": {
    "confidence_threshold": 0.7,
    "template_cache_size": 100,
    "max_template_variations": 3,
    "scale_factors": [0.9, 1.0, 1.1],
    "matching_method": "TM_CCOEFF_NORMED"
  },
  "smart_automation": {
    "method_priority": ["ocr", "template", "coordinates"],
    "max_retries": 2,
    "retry_delay": 0.5,
    "exponential_backoff": false,
    "screenshot_on_error": false,
    "performance_tracking": true
  },
  "performance": {
    "screenshot_region_optimization": true,
    "parallel_processing": true,
    "cache_successful_locations": true,
    "cache_duration_seconds": 600,
    "max_concurrent_operations": 3,
    "memory_cleanup_interval": 300
  },
  "debugging": {
    "debug_mode": false,
    "save_debug_images": false,
    "verbose_logging": false
  }
}
```

### 2. Adaptive Configuration

#### Dynamic Configuration Adjustment

```python
class AdaptiveConfigManager:
    def __init__(self):
        self.performance_history = []
        self.config_adjustments = {}
    
    def adjust_based_on_performance(self, performance_stats):
        """Automatically adjust configuration based on performance"""
        
        adjustments = {}
        
        # If OCR is slow, reduce preprocessing
        if performance_stats.get('average_ocr_time', 0) > 3.0:
            adjustments['ocr_settings.preprocessing.gaussian_blur_kernel'] = 1
            adjustments['ocr_settings.confidence_threshold'] = 0.5
        
        # If template matching is slow, reduce scale factors
        if performance_stats.get('average_template_time', 0) > 2.0:
            adjustments['template_matching.scale_factors'] = [1.0]
            adjustments['template_matching.max_template_variations'] = 1
        
        # If memory usage is high, reduce cache sizes
        memory_percent = psutil.virtual_memory().percent
        if memory_percent > 80:
            adjustments['template_matching.template_cache_size'] = 25
            adjustments['performance.cache_duration_seconds'] = 300
        
        return adjustments
    
    def apply_adjustments(self, adjustments):
        """Apply configuration adjustments"""
        for key, value in adjustments.items():
            self.set_nested_config(key, value)
            print(f"Adjusted {key} to {value}")
```

## Performance Monitoring

### 1. Real-time Performance Tracking

#### Performance Monitor Implementation

```python
import time
import threading
from collections import deque
from dataclasses import dataclass
from typing import Dict, List

@dataclass
class PerformanceMetric:
    timestamp: float
    operation: str
    duration: float
    success: bool
    memory_usage: float
    cpu_usage: float

class PerformanceMonitor:
    def __init__(self, history_size=1000):
        self.metrics = deque(maxlen=history_size)
        self.lock = threading.Lock()
        self.monitoring = True
        
        # Start background monitoring
        self.monitor_thread = threading.Thread(target=self._monitor_system)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
    
    def record_operation(self, operation: str, duration: float, success: bool):
        """Record an operation's performance"""
        
        metric = PerformanceMetric(
            timestamp=time.time(),
            operation=operation,
            duration=duration,
            success=success,
            memory_usage=psutil.virtual_memory().percent,
            cpu_usage=psutil.cpu_percent()
        )
        
        with self.lock:
            self.metrics.append(metric)
    
    def get_performance_summary(self, time_window=300):  # 5 minutes
        """Get performance summary for recent operations"""
        
        current_time = time.time()
        cutoff_time = current_time - time_window
        
        with self.lock:
            recent_metrics = [m for m in self.metrics if m.timestamp >= cutoff_time]
        
        if not recent_metrics:
            return {}
        
        # Calculate statistics
        total_operations = len(recent_metrics)
        successful_operations = sum(1 for m in recent_metrics if m.success)
        success_rate = successful_operations / total_operations
        
        durations = [m.duration for m in recent_metrics]
        avg_duration = sum(durations) / len(durations)
        max_duration = max(durations)
        
        memory_usage = [m.memory_usage for m in recent_metrics]
        avg_memory = sum(memory_usage) / len(memory_usage)
        
        cpu_usage = [m.cpu_usage for m in recent_metrics]
        avg_cpu = sum(cpu_usage) / len(cpu_usage)
        
        return {
            "time_window": time_window,
            "total_operations": total_operations,
            "success_rate": success_rate,
            "avg_duration": avg_duration,
            "max_duration": max_duration,
            "avg_memory_usage": avg_memory,
            "avg_cpu_usage": avg_cpu
        }
    
    def _monitor_system(self):
        """Background system monitoring"""
        while self.monitoring:
            try:
                # Check for performance issues
                summary = self.get_performance_summary(60)  # Last minute
                
                if summary:
                    # Alert on performance issues
                    if summary["success_rate"] < 0.8:
                        print(f"⚠️ Low success rate: {summary['success_rate']:.1%}")
                    
                    if summary["avg_duration"] > 5.0:
                        print(f"⚠️ Slow operations: {summary['avg_duration']:.1f}s average")
                    
                    if summary["avg_memory_usage"] > 80:
                        print(f"⚠️ High memory usage: {summary['avg_memory_usage']:.1f}%")
                
                time.sleep(30)  # Check every 30 seconds
                
            except Exception as e:
                print(f"Performance monitoring error: {e}")
                time.sleep(60)
```

### 2. Performance Benchmarking

#### Automated Benchmark Suite

```python
class PerformanceBenchmark:
    def __init__(self):
        self.results = {}
    
    def benchmark_ocr_performance(self, test_images, iterations=10):
        """Benchmark OCR performance with different configurations"""
        
        configs = [
            {"name": "fast", "confidence": 0.5, "blur": 1, "block": 9},
            {"name": "balanced", "confidence": 0.7, "blur": 3, "block": 11},
            {"name": "accurate", "confidence": 0.8, "blur": 5, "block": 13}
        ]
        
        results = {}
        
        for config in configs:
            config_results = []
            
            for image_path in test_images:
                image = cv2.imread(image_path)
                
                # Time OCR operation
                start_time = time.time()
                
                for _ in range(iterations):
                    # Apply configuration
                    ocr_service = OCRService()
                    ocr_service.config.update(config)
                    
                    # Perform OCR
                    matches = ocr_service.find_text(image, "test")
                
                avg_time = (time.time() - start_time) / iterations
                config_results.append(avg_time)
            
            results[config["name"]] = {
                "avg_time": sum(config_results) / len(config_results),
                "min_time": min(config_results),
                "max_time": max(config_results)
            }
        
        return results
    
    def benchmark_template_matching(self, screenshot, templates, iterations=10):
        """Benchmark template matching performance"""
        
        results = {}
        
        for template_name in templates:
            times = []
            
            for _ in range(iterations):
                start_time = time.time()
                
                template_service = TemplateService()
                result = template_service.find_template(screenshot, template_name)
                
                times.append(time.time() - start_time)
            
            results[template_name] = {
                "avg_time": sum(times) / len(times),
                "min_time": min(times),
                "max_time": max(times)
            }
        
        return results
```

## Optimization Recommendations

### 1. Quick Wins (Easy to Implement)

1. **Enable Region-of-Interest**: Limit OCR to specific screen areas
2. **Reduce Template Variations**: Use only essential template variations
3. **Lower Confidence Thresholds**: Balance accuracy vs. speed
4. **Enable Caching**: Cache successful element locations
5. **Disable Debug Features**: Turn off debug image saving in production

### 2. Medium Impact (Moderate Effort)

1. **Optimize Image Preprocessing**: Fine-tune preprocessing parameters
2. **Implement Parallel Processing**: Use threading for multiple operations
3. **Smart Template Management**: Implement intelligent template caching
4. **Adaptive Configuration**: Automatically adjust based on performance
5. **Memory Management**: Implement proper cleanup and garbage collection

### 3. High Impact (Significant Effort)

1. **Custom OCR Engine**: Implement specialized OCR for VBS UI
2. **Machine Learning Integration**: Use ML for element detection
3. **GPU Acceleration**: Leverage GPU for image processing
4. **Distributed Processing**: Scale across multiple machines
5. **Native Code Integration**: Use C++ extensions for critical paths

## Performance Testing

### Automated Performance Tests

```python
def run_performance_test_suite():
    """Run comprehensive performance test suite"""
    
    test_results = {}
    
    # Test 1: OCR Performance
    print("Testing OCR performance...")
    ocr_results = benchmark_ocr_performance()
    test_results["ocr"] = ocr_results
    
    # Test 2: Template Matching Performance
    print("Testing template matching performance...")
    template_results = benchmark_template_matching()
    test_results["template"] = template_results
    
    # Test 3: Memory Usage
    print("Testing memory usage...")
    memory_results = benchmark_memory_usage()
    test_results["memory"] = memory_results
    
    # Test 4: End-to-End Performance
    print("Testing end-to-end performance...")
    e2e_results = benchmark_end_to_end()
    test_results["end_to_end"] = e2e_results
    
    # Generate report
    generate_performance_report(test_results)
    
    return test_results

def generate_performance_report(results):
    """Generate comprehensive performance report"""
    
    report = {
        "timestamp": datetime.now().isoformat(),
        "system_info": {
            "cpu_count": os.cpu_count(),
            "memory_total": psutil.virtual_memory().total,
            "platform": platform.platform()
        },
        "test_results": results,
        "recommendations": generate_recommendations(results)
    }
    
    with open("performance_report.json", "w") as f:
        json.dump(report, f, indent=2)
    
    print("Performance report saved to performance_report.json")
```

This performance tuning guide provides comprehensive strategies for optimizing the VBS OpenCV modernization system. Regular performance monitoring and optimization based on actual usage patterns will ensure optimal system performance.