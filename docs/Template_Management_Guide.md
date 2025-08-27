# Template Management Guide

## Overview

This guide covers the creation, management, and optimization of template images for the VBS OpenCV modernization system. Templates are crucial for reliable UI element detection and should be carefully created and maintained.

## Template Basics

### What are Templates?

Templates are reference images of UI elements that the system uses to locate similar elements on the screen through image matching. They serve as visual "fingerprints" for buttons, menus, icons, and other interface components.

### When to Use Templates

Templates are ideal for:
- **Graphical elements**: Icons, buttons with images, logos
- **Consistent UI elements**: Elements that don't change text but have consistent visual appearance
- **OCR fallback**: When text recognition fails or is unreliable
- **Complex elements**: Multi-part UI components that are hard to describe with text

### Template vs. OCR Decision Matrix

| Element Type | Recommended Method | Reason |
|--------------|-------------------|---------|
| Text buttons | OCR | Text is more reliable than button appearance |
| Icon buttons | Template | No text to recognize |
| Menu items with text | OCR | Text-based detection is more flexible |
| Graphical indicators | Template | Visual patterns are key identifier |
| Mixed text/graphics | Both (Hybrid) | Use both methods for best reliability |

## Template Directory Structure

### Recommended Organization

```
vbs/templates/
├── navigation/                 # Phase 2 navigation elements
│   ├── arrow_button.png
│   ├── arrow_button_hover.png
│   ├── sales_menu.png
│   ├── pos_menu.png
│   └── wifi_menu.png
├── upload/                     # Phase 3 upload elements
│   ├── import_button.png
│   ├── import_button_pressed.png
│   ├── browse_button.png
│   ├── dropdown_arrow.png
│   └── progress_bar.png
├── reports/                    # Phase 4 report elements
│   ├── reports_menu.png
│   ├── print_button.png
│   ├── export_button.png
│   └── pdf_dialog.png
├── common/                     # Shared UI elements
│   ├── ok_button.png
│   ├── cancel_button.png
│   ├── close_button.png
│   └── minimize_button.png
└── dialogs/                    # Dialog box elements
    ├── file_dialog_open.png
    ├── save_dialog.png
    └── confirmation_dialog.png
```

### Naming Conventions

#### Standard Naming Pattern
```
{element_type}_{state}_{variation}.png
```

Examples:
- `import_button.png` - Basic import button
- `import_button_hover.png` - Import button in hover state
- `import_button_pressed.png` - Import button when pressed
- `import_button_disabled.png` - Import button when disabled

#### State Suffixes
- `_normal` - Default state (can be omitted)
- `_hover` - Mouse hover state
- `_pressed` - Button pressed state
- `_disabled` - Disabled/grayed out state
- `_selected` - Selected/active state
- `_focused` - Keyboard focus state

#### Variation Suffixes
- `_v1`, `_v2`, etc. - Different versions of the same element
- `_small`, `_large` - Size variations
- `_light`, `_dark` - Theme variations

## Template Creation Process

### 1. Screenshot Capture

#### Using Built-in Template Capture Tool

```python
from cv_services.template_capture import TemplateCapture

# Initialize capture tool
capture = TemplateCapture()

# Capture specific region
template_image = capture.capture_region(x=100, y=200, width=80, height=30)

# Save template
capture.save_template(template_image, "import_button", "upload")
```

#### Manual Screenshot Process

1. **Prepare the VBS Application**
   - Ensure VBS is in the correct state
   - Navigate to the screen containing the target element
   - Ensure element is clearly visible and unobstructed

2. **Capture Screenshot**
   - Use Windows Snipping Tool or similar
   - Capture the entire screen or VBS window
   - Save as PNG format for best quality

3. **Extract Template Region**
   - Open screenshot in image editor
   - Carefully select the target UI element
   - Include minimal surrounding area
   - Crop to exact element boundaries

### 2. Template Optimization

#### Image Quality Guidelines

```python
def optimize_template_image(image_path, output_path):
    """Optimize template image for better matching"""
    
    import cv2
    import numpy as np
    
    # Load image
    image = cv2.imread(image_path)
    
    # Convert to RGB if needed
    if len(image.shape) == 3:
        image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    
    # Resize if too large (max 200x200 for performance)
    height, width = image.shape[:2]
    if width > 200 or height > 200:
        scale = min(200/width, 200/height)
        new_width = int(width * scale)
        new_height = int(height * scale)
        image = cv2.resize(image, (new_width, new_height), interpolation=cv2.INTER_LANCZOS4)
    
    # Enhance contrast for better matching
    if len(image.shape) == 3:
        # Convert to LAB color space
        lab = cv2.cvtColor(image, cv2.COLOR_RGB2LAB)
        l, a, b = cv2.split(lab)
        
        # Apply CLAHE to L channel
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        l = clahe.apply(l)
        
        # Merge channels and convert back
        enhanced = cv2.merge([l, a, b])
        image = cv2.cvtColor(enhanced, cv2.COLOR_LAB2RGB)
    else:
        # Grayscale enhancement
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
        image = clahe.apply(image)
    
    # Save optimized template
    cv2.imwrite(output_path, cv2.cvtColor(image, cv2.COLOR_RGB2BGR))
    
    return image
```

#### Template Quality Checklist

- ✅ **Clear and Sharp**: No blur or distortion
- ✅ **Appropriate Size**: 20x20 to 200x200 pixels optimal
- ✅ **Minimal Background**: Only the target element, minimal surrounding area
- ✅ **Consistent Lighting**: Avoid shadows or unusual lighting conditions
- ✅ **Standard State**: Capture in normal/default state unless specifically needed
- ✅ **High Contrast**: Element should stand out from background
- ✅ **No Overlapping Elements**: Avoid other UI elements in the template
- ✅ **Stable Elements**: Avoid elements that change frequently (like timestamps)

### 3. Template Validation

#### Automated Template Testing

```python
class TemplateValidator:
    def __init__(self):
        self.template_service = TemplateService()
    
    def validate_template(self, template_name, test_screenshots):
        """Validate template against multiple test screenshots"""
        
        results = {
            "template_name": template_name,
            "test_results": [],
            "success_rate": 0,
            "avg_confidence": 0,
            "recommendations": []
        }
        
        successful_matches = 0
        total_confidence = 0
        
        for screenshot_path in test_screenshots:
            screenshot = cv2.imread(screenshot_path)
            
            # Test template matching
            match_result = self.template_service.find_template(screenshot, template_name)
            
            test_result = {
                "screenshot": screenshot_path,
                "found": match_result.success,
                "confidence": match_result.matches[0].confidence if match_result.matches else 0,
                "location": match_result.matches[0].location if match_result.matches else None
            }
            
            results["test_results"].append(test_result)
            
            if match_result.success:
                successful_matches += 1
                total_confidence += match_result.matches[0].confidence
        
        # Calculate statistics
        results["success_rate"] = successful_matches / len(test_screenshots)
        results["avg_confidence"] = total_confidence / successful_matches if successful_matches > 0 else 0
        
        # Generate recommendations
        if results["success_rate"] < 0.8:
            results["recommendations"].append("Low success rate - consider creating template variations")
        
        if results["avg_confidence"] < 0.8:
            results["recommendations"].append("Low confidence - template may be too generic or poor quality")
        
        return results
    
    def batch_validate_templates(self, template_directory, test_screenshots_directory):
        """Validate all templates in a directory"""
        
        import os
        import glob
        
        template_files = glob.glob(os.path.join(template_directory, "*.png"))
        test_screenshots = glob.glob(os.path.join(test_screenshots_directory, "*.png"))
        
        validation_results = []
        
        for template_file in template_files:
            template_name = os.path.splitext(os.path.basename(template_file))[0]
            result = self.validate_template(template_name, test_screenshots)
            validation_results.append(result)
        
        return validation_results
```

## Template Management Tools

### 1. Template Capture Utility

#### Interactive Template Capture

```python
import tkinter as tk
from tkinter import filedialog, messagebox
import cv2
import numpy as np
from PIL import Image, ImageTk

class TemplateCaptureTool:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("VBS Template Capture Tool")
        self.root.geometry("800x600")
        
        self.screenshot = None
        self.selection_start = None
        self.selection_end = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        
        # Control frame
        control_frame = tk.Frame(self.root)
        control_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        tk.Button(control_frame, text="Capture Screen", command=self.capture_screen).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Load Image", command=self.load_image).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Save Template", command=self.save_template).pack(side=tk.LEFT, padx=5)
        
        # Canvas for image display
        self.canvas = tk.Canvas(self.root, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Bind mouse events for selection
        self.canvas.bind("<Button-1>", self.start_selection)
        self.canvas.bind("<B1-Motion>", self.update_selection)
        self.canvas.bind("<ButtonRelease-1>", self.end_selection)
    
    def capture_screen(self):
        """Capture current screen"""
        import pyautogui
        
        # Hide window temporarily
        self.root.withdraw()
        
        # Wait a moment for window to hide
        self.root.after(500, self._do_capture)
    
    def _do_capture(self):
        """Perform the actual screen capture"""
        import pyautogui
        
        screenshot = pyautogui.screenshot()
        self.screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)
        
        # Show window again
        self.root.deiconify()
        
        # Display screenshot
        self.display_image(self.screenshot)
    
    def display_image(self, image):
        """Display image on canvas"""
        
        # Convert to RGB for display
        if len(image.shape) == 3:
            display_image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        else:
            display_image = image
        
        # Resize to fit canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width > 1 and canvas_height > 1:
            height, width = display_image.shape[:2]
            scale = min(canvas_width/width, canvas_height/height)
            
            new_width = int(width * scale)
            new_height = int(height * scale)
            
            display_image = cv2.resize(display_image, (new_width, new_height))
        
        # Convert to PhotoImage
        pil_image = Image.fromarray(display_image)
        self.photo = ImageTk.PhotoImage(pil_image)
        
        # Clear canvas and display image
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
    
    def start_selection(self, event):
        """Start selection rectangle"""
        self.selection_start = (event.x, event.y)
        self.canvas.delete("selection")
    
    def update_selection(self, event):
        """Update selection rectangle"""
        if self.selection_start:
            self.canvas.delete("selection")
            self.canvas.create_rectangle(
                self.selection_start[0], self.selection_start[1],
                event.x, event.y,
                outline="red", width=2, tags="selection"
            )
    
    def end_selection(self, event):
        """End selection rectangle"""
        self.selection_end = (event.x, event.y)
    
    def save_template(self):
        """Save selected region as template"""
        if not self.screenshot or not self.selection_start or not self.selection_end:
            messagebox.showerror("Error", "Please capture screen and select region first")
            return
        
        # Get selection coordinates
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        # Ensure correct order
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        
        # Calculate scale factor
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_height, img_width = self.screenshot.shape[:2]
        
        scale_x = img_width / canvas_width
        scale_y = img_height / canvas_height
        
        # Convert to image coordinates
        img_x1 = int(x1 * scale_x)
        img_y1 = int(y1 * scale_y)
        img_x2 = int(x2 * scale_x)
        img_y2 = int(y2 * scale_y)
        
        # Extract template region
        template = self.screenshot[img_y1:img_y2, img_x1:img_x2]
        
        if template.size == 0:
            messagebox.showerror("Error", "Invalid selection")
            return
        
        # Ask for template name and category
        template_name = tk.simpledialog.askstring("Template Name", "Enter template name:")
        if not template_name:
            return
        
        category = tk.simpledialog.askstring("Category", "Enter category (navigation/upload/reports/common):")
        if not category:
            category = "common"
        
        # Save template
        template_dir = f"vbs/templates/{category}"
        os.makedirs(template_dir, exist_ok=True)
        
        template_path = os.path.join(template_dir, f"{template_name}.png")
        cv2.imwrite(template_path, template)
        
        messagebox.showinfo("Success", f"Template saved to {template_path}")
    
    def run(self):
        """Run the template capture tool"""
        self.root.mainloop()

# Usage
if __name__ == "__main__":
    tool = TemplateCaptureTool()
    tool.run()
```

### 2. Template Management System

#### Template Database Manager

```python
import json
import os
import hashlib
from datetime import datetime
from typing import Dict, List, Optional

class TemplateDatabase:
    def __init__(self, templates_dir="vbs/templates"):
        self.templates_dir = templates_dir
        self.db_file = os.path.join(templates_dir, "template_database.json")
        self.database = self.load_database()
    
    def load_database(self) -> Dict:
        """Load template database from file"""
        if os.path.exists(self.db_file):
            try:
                with open(self.db_file, 'r') as f:
                    return json.load(f)
            except:
                pass
        
        return {
            "version": "1.0",
            "created": datetime.now().isoformat(),
            "templates": {}
        }
    
    def save_database(self):
        """Save template database to file"""
        self.database["last_updated"] = datetime.now().isoformat()
        
        os.makedirs(os.path.dirname(self.db_file), exist_ok=True)
        with open(self.db_file, 'w') as f:
            json.dump(self.database, f, indent=2)
    
    def add_template(self, template_name: str, category: str, 
                    description: str = "", tags: List[str] = None,
                    confidence_threshold: float = 0.8):
        """Add template to database"""
        
        template_path = os.path.join(self.templates_dir, category, f"{template_name}.png")
        
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
        
        # Calculate file hash for integrity checking
        with open(template_path, 'rb') as f:
            file_hash = hashlib.md5(f.read()).hexdigest()
        
        # Get image dimensions
        import cv2
        image = cv2.imread(template_path)
        height, width = image.shape[:2]
        
        template_info = {
            "name": template_name,
            "category": category,
            "description": description,
            "tags": tags or [],
            "file_path": template_path,
            "file_hash": file_hash,
            "dimensions": {"width": width, "height": height},
            "confidence_threshold": confidence_threshold,
            "created": datetime.now().isoformat(),
            "usage_count": 0,
            "success_rate": 0.0,
            "variations": []
        }
        
        self.database["templates"][template_name] = template_info
        self.save_database()
    
    def add_template_variation(self, base_template: str, variation_name: str, 
                             variation_type: str = "state"):
        """Add a variation to an existing template"""
        
        if base_template not in self.database["templates"]:
            raise ValueError(f"Base template '{base_template}' not found")
        
        variation_info = {
            "name": variation_name,
            "type": variation_type,  # state, size, theme, etc.
            "created": datetime.now().isoformat()
        }
        
        self.database["templates"][base_template]["variations"].append(variation_info)
        self.save_database()
    
    def update_template_stats(self, template_name: str, success: bool):
        """Update template usage statistics"""
        
        if template_name in self.database["templates"]:
            template = self.database["templates"][template_name]
            template["usage_count"] += 1
            
            # Update success rate (simple moving average)
            current_rate = template["success_rate"]
            usage_count = template["usage_count"]
            
            if success:
                template["success_rate"] = (current_rate * (usage_count - 1) + 1.0) / usage_count
            else:
                template["success_rate"] = (current_rate * (usage_count - 1)) / usage_count
            
            self.save_database()
    
    def get_template_info(self, template_name: str) -> Optional[Dict]:
        """Get information about a template"""
        return self.database["templates"].get(template_name)
    
    def list_templates(self, category: str = None, tag: str = None) -> List[Dict]:
        """List templates with optional filtering"""
        
        templates = []
        
        for name, info in self.database["templates"].items():
            # Filter by category
            if category and info["category"] != category:
                continue
            
            # Filter by tag
            if tag and tag not in info["tags"]:
                continue
            
            templates.append(info)
        
        return templates
    
    def validate_templates(self) -> Dict[str, List[str]]:
        """Validate all templates in database"""
        
        results = {
            "valid": [],
            "missing_files": [],
            "hash_mismatches": [],
            "errors": []
        }
        
        for name, info in self.database["templates"].items():
            try:
                file_path = info["file_path"]
                
                # Check if file exists
                if not os.path.exists(file_path):
                    results["missing_files"].append(name)
                    continue
                
                # Check file hash
                with open(file_path, 'rb') as f:
                    current_hash = hashlib.md5(f.read()).hexdigest()
                
                if current_hash != info["file_hash"]:
                    results["hash_mismatches"].append(name)
                    continue
                
                results["valid"].append(name)
                
            except Exception as e:
                results["errors"].append(f"{name}: {str(e)}")
        
        return results
    
    def cleanup_unused_templates(self, dry_run: bool = True) -> List[str]:
        """Remove templates with very low success rates"""
        
        to_remove = []
        
        for name, info in self.database["templates"].items():
            # Remove templates with success rate < 20% and usage > 10
            if info["usage_count"] > 10 and info["success_rate"] < 0.2:
                to_remove.append(name)
        
        if not dry_run:
            for name in to_remove:
                template_info = self.database["templates"][name]
                
                # Move file to archive
                archive_dir = os.path.join(self.templates_dir, "archive")
                os.makedirs(archive_dir, exist_ok=True)
                
                old_path = template_info["file_path"]
                new_path = os.path.join(archive_dir, os.path.basename(old_path))
                
                if os.path.exists(old_path):
                    os.rename(old_path, new_path)
                
                # Remove from database
                del self.database["templates"][name]
            
            self.save_database()
        
        return to_remove
```

## Template Optimization Strategies

### 1. Multi-Scale Templates

Create templates at different scales for better matching:

```python
def create_multi_scale_templates(original_template_path, output_dir):
    """Create multiple scale variations of a template"""
    
    import cv2
    import os
    
    # Load original template
    template = cv2.imread(original_template_path)
    base_name = os.path.splitext(os.path.basename(original_template_path))[0]
    
    # Define scale factors
    scales = [0.8, 0.9, 1.0, 1.1, 1.2]
    
    for scale in scales:
        if scale == 1.0:
            # Save original
            output_path = os.path.join(output_dir, f"{base_name}.png")
            cv2.imwrite(output_path, template)
        else:
            # Create scaled version
            height, width = template.shape[:2]
            new_height = int(height * scale)
            new_width = int(width * scale)
            
            scaled = cv2.resize(template, (new_width, new_height), interpolation=cv2.INTER_LANCZOS4)
            
            # Save scaled version
            scale_str = str(scale).replace('.', '_')
            output_path = os.path.join(output_dir, f"{base_name}_scale_{scale_str}.png")
            cv2.imwrite(output_path, scaled)
```

### 2. Template Preprocessing

Optimize templates for better matching:

```python
def preprocess_template_for_matching(template_path, output_path):
    """Preprocess template for optimal matching performance"""
    
    import cv2
    import numpy as np
    
    # Load template
    template = cv2.imread(template_path)
    
    # Convert to grayscale if needed
    if len(template.shape) == 3:
        gray = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
    else:
        gray = template
    
    # Apply histogram equalization for better contrast
    equalized = cv2.equalizeHist(gray)
    
    # Apply slight Gaussian blur to reduce noise
    blurred = cv2.GaussianBlur(equalized, (3, 3), 0)
    
    # Save preprocessed template
    cv2.imwrite(output_path, blurred)
    
    return blurred
```

## Best Practices

### 1. Template Creation Best Practices

1. **Capture in Standard State**: Always capture UI elements in their normal, default state
2. **Minimal Background**: Include only the target element with minimal surrounding area
3. **Consistent Lighting**: Avoid unusual lighting conditions or shadows
4. **High Quality**: Use high-resolution screenshots and avoid compression artifacts
5. **Multiple Variations**: Create variations for different states (hover, pressed, disabled)
6. **Regular Updates**: Update templates when UI changes occur

### 2. Template Organization Best Practices

1. **Logical Categorization**: Group templates by functionality or VBS phase
2. **Consistent Naming**: Use clear, descriptive names with consistent conventions
3. **Version Control**: Keep track of template versions and changes
4. **Documentation**: Document the purpose and usage of each template
5. **Regular Cleanup**: Remove unused or poorly performing templates

### 3. Template Maintenance Best Practices

1. **Regular Validation**: Periodically validate templates against current UI
2. **Performance Monitoring**: Track template matching success rates
3. **Automated Testing**: Include templates in automated test suites
4. **Backup Strategy**: Maintain backups of working template sets
5. **Change Management**: Document template changes and their impact

## Troubleshooting Template Issues

### Common Template Problems

1. **Template Not Found**
   - Check file path and naming
   - Verify template exists in correct directory
   - Check file permissions

2. **Low Matching Confidence**
   - Update template with current UI screenshot
   - Create multiple template variations
   - Adjust confidence threshold
   - Check for UI scaling issues

3. **False Positives**
   - Make template more specific
   - Increase confidence threshold
   - Add negative templates (what not to match)
   - Use region-of-interest to limit search area

4. **Performance Issues**
   - Reduce template size
   - Limit number of template variations
   - Use grayscale templates
   - Implement template caching

### Template Debugging Tools

```python
def debug_template_matching(screenshot_path, template_path, output_path):
    """Debug template matching by visualizing results"""
    
    import cv2
    import numpy as np
    
    # Load images
    screenshot = cv2.imread(screenshot_path)
    template = cv2.imread(template_path)
    
    # Convert to grayscale
    gray_screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
    gray_template = cv2.cvtColor(template, cv2.COLOR_BGR2GRAY)
    
    # Perform template matching
    result = cv2.matchTemplate(gray_screenshot, gray_template, cv2.TM_CCOEFF_NORMED)
    
    # Find all matches above threshold
    threshold = 0.7
    locations = np.where(result >= threshold)
    
    # Draw rectangles around matches
    debug_image = screenshot.copy()
    template_height, template_width = gray_template.shape
    
    for pt in zip(*locations[::-1]):
        cv2.rectangle(debug_image, pt, (pt[0] + template_width, pt[1] + template_height), (0, 255, 0), 2)
        
        # Add confidence score
        confidence = result[pt[1], pt[0]]
        cv2.putText(debug_image, f"{confidence:.2f}", (pt[0], pt[1] - 10), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
    
    # Save debug image
    cv2.imwrite(output_path, debug_image)
    
    print(f"Found {len(locations[0])} matches above threshold {threshold}")
    print(f"Debug image saved to {output_path}")
```

This comprehensive template management guide provides all the tools and knowledge needed to create, maintain, and optimize templates for the VBS OpenCV modernization system.