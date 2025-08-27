# VBS UI Templates

This directory contains template images for VBS UI element detection.

## Directory Structure

```
vbs/templates/
├── phase2_navigation/
│   ├── arrow_button.png
│   ├── sales_distribution.png
│   ├── pos_menu.png
│   └── wifi_registration.png
├── phase3_upload/
│   ├── import_ehc_checkbox.png
│   ├── three_dots_button.png
│   ├── dropdown_button.png
│   └── import_button.png
├── phase4_report/
│   ├── reports_menu.png
│   ├── pos_reports.png
│   ├── print_button.png
│   └── export_button.png
└── common/
    ├── new_button.png
    ├── ok_button.png
    ├── cancel_button.png
    └── update_button.png
```

## Template Guidelines

1. **Image Format**: PNG format preferred for transparency support
2. **Resolution**: Capture at actual size (100% zoom)
3. **Quality**: High quality, clear images without compression artifacts
4. **Variations**: Create multiple variations for different UI states
5. **Naming**: Use descriptive names with phase prefix

## Template Creation

Use the template capture utility:
```python
from vbs.cv_services import TemplateService
template_service = TemplateService()
template_service.capture_template("element_name", region=(x, y, width, height))
```

## Template Validation

Templates are automatically validated for:
- Image quality and clarity
- Uniqueness (not too similar to existing templates)
- Appropriate size and aspect ratio
- Sufficient contrast for matching