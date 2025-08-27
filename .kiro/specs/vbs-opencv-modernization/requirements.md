# VBS Automation Modernization with OpenCV Computer Vision

## Introduction

This feature modernizes the existing VBS automation system by integrating OpenCV and Tesseract OCR for robust computer vision-based automation. The current system relies on fragile coordinate-based clicking which fails when UI elements move or change. The new system will use advanced computer vision techniques including OCR, template matching, and intelligent element detection to create a more reliable, adaptive, and faster automation solution for VBS phases 2, 3, and 4, with a comprehensive batch coordination system.

## Requirements

### Requirement 1: OpenCV Computer Vision Integration

**User Story:** As a system administrator, I want the VBS automation to use OpenCV computer vision instead of hardcoded coordinates, so that the automation is more reliable and adapts to UI changes automatically.

#### Acceptance Criteria

1. WHEN the system initializes THEN it SHALL load OpenCV and Tesseract libraries successfully
2. WHEN the system detects UI elements THEN it SHALL use template matching and OCR instead of fixed coordinates
3. WHEN UI elements move or change appearance THEN the system SHALL still locate them with 90%+ accuracy
4. IF computer vision fails THEN the system SHALL fallback to coordinate-based automation with detailed logging
5. WHEN using computer vision THEN the system SHALL complete element detection within 2 seconds per element

### Requirement 2: Enhanced VBS Phase 2 Navigation

**User Story:** As a system operator, I want Phase 2 navigation to use computer vision for menu detection, so that it can reliably navigate Sales & Distribution → POS → WiFi User Registration even when UI layout changes.

#### Acceptance Criteria

1. WHEN Phase 2 starts THEN it SHALL use OCR to detect "Sales and Distribution" menu text instead of coordinates (132, 165)
2. WHEN navigating to POS THEN it SHALL use template matching to find POS menu item instead of coordinates (159,598)
3. WHEN reaching WiFi User Registration THEN it SHALL use keyboard navigation (3 TABS + ENTER) as specified in vbsphases.txt
4. WHEN New button is needed THEN it SHALL use 2 LEFT ARROW KEYS + ENTER as specified
5. IF computer vision fails THEN the system SHALL fallback to original coordinate-based method

### Requirement 3: Enhanced VBS Phase 3 Upload Process

**User Story:** As a system operator, I want Phase 3 upload to use computer vision for form detection, so that it can reliably handle Excel import and update processes even when dialog layouts change.

#### Acceptance Criteria

1. WHEN Phase 3 starts THEN it SHALL use OCR to detect "Import EHC" checkbox text instead of coordinates (1194,692)
2. WHEN file selection is needed THEN it SHALL use template matching to find the 3 dots button instead of coordinates (785,658)
3. WHEN dropdown selection is needed THEN it SHALL use OCR to detect "Sheet1" text instead of coordinates (432,715)
4. WHEN monitoring update progress THEN it SHALL use visual detection for completion instead of only audio cues
5. WHEN update completes THEN it SHALL detect success popup using OCR before pressing ENTER

### Requirement 4: Enhanced VBS Phase 4 Report Generation

**User Story:** As a system operator, I want Phase 4 report generation to use computer vision for PDF interface detection, so that it can reliably generate and export reports even when interface elements move.

#### Acceptance Criteria

1. WHEN Phase 4 starts THEN it SHALL use OCR to detect "Reports" menu text instead of coordinates (157,646)
2. WHEN navigating to POS reports THEN it SHALL use template matching to find POS within reports instead of coordinates (187,934)
3. WHEN entering dates THEN it SHALL use OCR to detect date field labels and validate input format
4. WHEN PDF generation starts THEN it SHALL use template matching to detect Print button instead of coordinates (114,110)
5. WHEN exporting PDF THEN it SHALL use OCR to detect file dialog elements and navigate intelligently

### Requirement 5: Multi-Method Fallback System

**User Story:** As a system operator, I want the automation to try multiple detection methods for each action, so that the system continues working even when one method fails.

#### Acceptance Criteria

1. WHEN an automation action is required THEN the system SHALL attempt methods in priority order: OCR → Template Matching → Coordinate-based
2. WHEN a method fails THEN the system SHALL automatically try the next method without manual intervention
3. WHEN all methods fail THEN the system SHALL log detailed error information and retry up to 3 times with screenshots
4. WHEN a method succeeds THEN the system SHALL record which method worked for performance optimization
5. WHEN the system learns successful patterns THEN it SHALL prioritize working methods for similar elements

### Requirement 6: Robust Batch Coordination System

**User Story:** As a system administrator, I want a comprehensive batch file system that coordinates all VBS phases, so that the entire workflow runs automatically with proper error handling and recovery.

#### Acceptance Criteria

1. WHEN the batch system starts THEN it SHALL execute phases in sequence: Phase 2 → Phase 3 → Phase 4
2. WHEN a phase fails THEN the system SHALL attempt recovery and retry up to 3 times before stopping
3. WHEN Phase 3 closes the application THEN the system SHALL automatically restart VBS for Phase 4
4. WHEN all phases complete THEN the system SHALL generate a comprehensive execution report
5. WHEN errors occur THEN the system SHALL capture screenshots and detailed logs for troubleshooting

### Requirement 7: OCR Text Recognition Service

**User Story:** As an automation engineer, I want the system to read text from the VBS application using OCR, so that it can intelligently interact with menus, buttons, and form fields.

#### Acceptance Criteria

1. WHEN the system encounters menu items THEN it SHALL use Tesseract OCR to read menu text with 85%+ confidence
2. WHEN the system needs to validate form fields THEN it SHALL read field labels and values using OCR
3. WHEN the system encounters buttons THEN it SHALL identify them by reading button text instead of coordinates
4. IF OCR confidence is below 80% THEN the system SHALL use template matching as fallback
5. WHEN processing text THEN the system SHALL handle VBS application fonts and text styles accurately

### Requirement 8: Template Matching System

**User Story:** As a developer, I want to manage UI element templates easily, so that the system can recognize buttons, icons, and UI elements without relying on coordinates.

#### Acceptance Criteria

1. WHEN new UI elements are encountered THEN the system SHALL allow capturing new templates from Images/phase2, Images/phase3, Images/phase4 directories
2. WHEN templates become outdated THEN the system SHALL detect low match confidence and suggest updates
3. WHEN multiple template variations exist THEN the system SHALL store and try all variations with confidence scoring
4. WHEN templates are matched THEN the system SHALL achieve 80%+ confidence before clicking
5. WHEN managing templates THEN the system SHALL provide utilities for template capture, validation, and optimization

### Requirement 9: Performance Optimization and Monitoring

**User Story:** As a system administrator, I want the modernized automation to be faster and more reliable than the current system, so that daily operations complete efficiently.

#### Acceptance Criteria

1. WHEN using computer vision THEN element detection SHALL complete within 2 seconds per element
2. WHEN using OCR THEN text recognition SHALL complete within 1 second per text area
3. WHEN the system processes a complete phase THEN it SHALL be at least 20% faster than coordinate-based automation
4. WHEN the system runs continuously THEN memory usage SHALL not exceed 500MB
5. WHEN processing images THEN the system SHALL optimize image preprocessing for speed and accuracy

### Requirement 10: Enhanced Error Handling and Recovery

**User Story:** As a system operator, I want comprehensive error handling with automatic recovery, so that I can quickly resolve issues and minimize downtime.

#### Acceptance Criteria

1. WHEN computer vision fails THEN the system SHALL capture screenshots and log detailed error context with timestamps
2. WHEN OCR produces low confidence results THEN the system SHALL save the problematic image for analysis and retry with different parameters
3. WHEN template matching fails THEN the system SHALL try alternative templates and log match confidence scores
4. WHEN any method fails THEN the system SHALL provide actionable error messages with suggested fixes
5. WHEN errors occur repeatedly THEN the system SHALL automatically adjust detection parameters and notify administrators

### Requirement 11: Configuration and Calibration System

**User Story:** As a system administrator, I want to configure computer vision parameters easily, so that the system can be tuned for optimal performance in different environments.

#### Acceptance Criteria

1. WHEN the system starts THEN it SHALL load computer vision configuration from config/cv_config.json
2. WHEN OCR accuracy is poor THEN administrators SHALL be able to adjust Tesseract parameters through configuration
3. WHEN template matching fails THEN administrators SHALL be able to update template images and thresholds
4. WHEN UI scaling changes THEN the system SHALL automatically detect and adjust for DPI changes
5. WHEN configuration changes THEN the system SHALL validate settings and provide immediate feedback

### Requirement 12: Integration with Existing VBS Workflow

**User Story:** As a system integrator, I want the computer vision enhancements to work seamlessly with existing VBS phases, so that the upgrade maintains compatibility with current workflows.

#### Acceptance Criteria

1. WHEN Phase 2 executes THEN it SHALL follow the exact navigation sequence from vbsphases.txt with CV enhancements
2. WHEN Phase 3 executes THEN it SHALL handle the complete upload workflow from vbsupdate.txt with visual monitoring
3. WHEN Phase 4 executes THEN it SHALL generate PDF reports following the exact process from vbsreport.txt
4. WHEN any phase completes THEN it SHALL return the same result format as the original implementation
5. WHEN the system encounters legacy coordinate requirements THEN it SHALL maintain backward compatibility