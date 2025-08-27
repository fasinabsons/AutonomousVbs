# Technology Stack & Build System

## Core Technologies

### Python Ecosystem
- **Python 3.11+**: Primary development language
- **Virtual Environment**: Recommended for dependency isolation

### Key Libraries & Frameworks
- **Selenium 4.15.2**: Web automation and browser control
- **undetected-chromedriver 3.5.4**: Chrome automation without detection
- **pandas 2.1.4**: Data processing and CSV manipulation
- **xlwt 1.3.0**: Excel file generation (.xls format for VBS compatibility)
- **pywin32 306**: Windows service integration and system APIs
- **schedule 1.2.0**: Task scheduling and automation timing
- **beautifulsoup4 4.12.2**: HTML parsing and web scraping
- **requests 2.31.0**: HTTP client for API interactions

### Windows Integration
- **Windows Services**: Background service operation using pywin32
- **Task Scheduler**: Secondary startup mechanism
- **Registry Integration**: System startup configuration
- **PowerShell/Batch**: Installation and management scripts

### Browser Automation
- **Chrome WebDriver**: Primary browser for web automation
- **Selenium Grid**: Distributed automation support
- **Headless Mode**: Background operation without UI

## Build & Development Commands

### Environment Setup
```bash
# Install dependencies
pip install -r requirements.txt

# Install Windows service dependencies (if needed)
pip install pywin32
python Scripts/pywin32_postinstall.py -install
```

### Service Management
```bash
# Install service
install_service.bat

# Manual service operations
python service_manager.py install
python service_manager.py start
python service_manager.py stop
python service_manager.py status

# Uninstall service
uninstall_service.bat
```

### Testing & Validation
```bash
# Component testing
python test_service_components.py
python test_service_integration.py

# Individual module testing
python test_csv_download.py
python test_vbs_integration.py

# Excel generation testing
python excel/excel_generator.py
```

### Manual Execution
```bash
# Run complete workflow
python orchestrator.py workflow

# Start continuous service
python orchestrator.py service

# Individual components
python wifi/csv_downloader.py
python vbs/vbs_automator.py --phase all
```

### Debugging & Monitoring
```bash
# View service logs
type EHC_Logs\windows_service.log
type EHC_Logs\orchestrator.log

# Health monitoring
python service_monitor.py

# Configuration validation
python utils/config_manager.py
```

## File Format Requirements

### Excel Compatibility
- **Format**: .xls (Excel 97-2003) for VBS compatibility
- **Encoding**: UTF-8 with BOM for special characters
- **Headers**: Exact column names required for VBS automation

### CSV Processing
- **Input**: UTF-8 encoded CSV files from Ruckus controller
- **Output**: Processed and merged data in daily folders
- **Validation**: Row count and data integrity checks

### Configuration
- **Format**: JSON for all configuration files
- **Validation**: Schema validation for critical settings
- **Backup**: Automatic configuration backup before changes