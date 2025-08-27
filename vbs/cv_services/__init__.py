"""
Computer Vision Services for VBS Automation
Provides OCR, template matching, and smart automation capabilities
"""

__version__ = "1.0.0"
__author__ = "VBS Automation Team"

# Import main services
from .ocr_service import OCRService
from .template_service import TemplateService
from .smart_engine import SmartAutomationEngine
from .element_detector import ElementDetector
from .error_handler import CVErrorHandler

__all__ = [
    'OCRService',
    'TemplateService', 
    'SmartAutomationEngine',
    'ElementDetector',
    'CVErrorHandler'
]