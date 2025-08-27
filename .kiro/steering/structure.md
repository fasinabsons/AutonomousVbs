# Project Structure & Organization

## Root Directory Layout

```
MoonFlower/
├── orchestrator.py              # Master orchestrator - main coordination logic
├── windows_service.py           # Windows service wrapper
├── service_manager.py           # Service installation and management
├── service_monitor.py           # Health monitoring and alerts
├── pc_restart_manager.py        # PC restart automation
├── requirements.txt             # Python dependencies
├── service_config.json          # Service configuration
├── install_service.bat          # Service installation script
├── uninstall_service.bat        # Service removal script
└── *.py                        # Various test and utility scripts
```

## Core Module Structure

### `/utils/` - Shared Utilities
- `config_manager.py` - Configuration loading and validation
- `file_manager.py` - File system operations and daily folder management
- `__pycache__/` - Python bytecode cache

### `/wifi/` - CSV Download Module
- `csv_downloader.py` - Ruckus controller automation and CSV downloads
- `__pycache__/` - Python bytecode cache

### `/excel/` - Excel Processing Module
- `excel_generator.py` - CSV to Excel conversion with VBS compatibility
- `__pycache__/` - Python bytecode cache

### `/vbs/` - VBS Automation Module
- `vbs_automator.py` - Main VBS automation entry point
- `vbs_core.py` - Core VBS automation logic
- `vbs_phase1_login.py` - Login and authentication phase
- `vbs_phase2_navigation.py` - Navigation and menu handling
- `vbs_phase3_upload.py` - File upload automation
- `vbs_phase4_report.py` - Report generation phase
- `vbs_audio_detector.py` - Audio cue detection for automation
- `vbs_error_handler.py` - Error handling and recovery
- `test_*.py` - VBS module testing scripts
- `__pycache__/` - Python bytecode cache

### `/email/` - Email Notification Module
- `email_delivery.py` - Email notifications and report delivery

### `/config/` - Configuration Files
- `settings.json` - Main application configuration

## Data Directory Structure

### Daily Data Organization (DDMmm format)
```
EHC_Data/17jul/                  # CSV files from Ruckus controller
EHC_Data_Merge/17jul/            # Processed Excel files for VBS
EHC_Data_Pdf/17jul/              # Generated PDF reports
EHC_Logs/17jul/                  # Daily operation logs
```

### Log Files
- `EHC_Logs/orchestrator.log` - Main orchestrator logs
- `EHC_Logs/windows_service.log` - Windows service logs
- `EHC_Logs/vbs_automator_YYYYMMDD.log` - VBS automation logs
- `service_install.log` - Service installation logs

## Naming Conventions

### Python Files
- **Snake case**: `file_manager.py`, `csv_downloader.py`
- **Descriptive names**: Clearly indicate module purpose
- **Phase naming**: VBS phases use `vbs_phaseN_description.py` pattern

### Configuration Files
- **JSON format**: All configuration in `.json` files
- **Descriptive names**: `service_config.json`, `settings.json`

### Data Files
- **Date folders**: `DDMmm` format (e.g., `17jul`, `04aug`)
- **Excel files**: `EHC_Upload_Mac_DDMMYYYY.xls` format
- **Log files**: Include date stamps for daily rotation

### Test Files
- **Prefix**: All test files start with `test_`
- **Module alignment**: `test_service_components.py` tests service components
- **Integration tests**: `test_service_integration.py` for cross-module testing

## Architecture Patterns

### Service Layer Pattern
- **orchestrator.py**: Central coordination and workflow management
- **service_manager.py**: Service lifecycle management
- **service_monitor.py**: Health monitoring and alerting

### Module Separation
- **Single responsibility**: Each module handles one specific domain
- **Loose coupling**: Modules communicate through well-defined interfaces
- **Dependency injection**: Configuration and file managers passed to modules

### Error Handling Strategy
- **Graceful degradation**: System continues with reduced functionality
- **Retry mechanisms**: Exponential backoff for transient failures
- **Comprehensive logging**: All errors logged with context and stack traces

### Configuration Management
- **Centralized**: All configuration through ConfigManager
- **Hierarchical**: Support for nested configuration keys with dot notation
- **Validation**: Configuration validation at startup and runtime

## File Organization Rules

### Import Structure
- **Relative imports**: Use relative imports within modules
- **Absolute imports**: Use absolute imports for cross-module dependencies
- **Path management**: Add parent directories to sys.path when needed

### Logging Standards
- **Module-level loggers**: Each module gets its own logger
- **Consistent formatting**: Standardized log message format
- **Log rotation**: Automatic rotation based on file size (50MB limit)

### Documentation
- **Docstrings**: All classes and functions have descriptive docstrings
- **Type hints**: Use typing module for function signatures
- **Comments**: Inline comments for complex logic and business rules