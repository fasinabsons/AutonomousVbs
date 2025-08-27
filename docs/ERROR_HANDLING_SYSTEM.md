# MoonFlower Error Handling and Recovery System

## Overview

The MoonFlower WiFi Automation system includes a comprehensive error handling and recovery framework designed to ensure 365-day continuous operation. The system automatically categorizes errors, attempts recovery, and provides detailed logging and alerting capabilities.

## Architecture

### Core Components

1. **ErrorHandler** (`utils/error_handler.py`)
   - Central error management and coordination
   - Automatic error categorization and severity detection
   - Recovery strategy execution
   - Error tracking and pattern analysis
   - Alert generation and processing

2. **RecoveryManager** (`utils/recovery_manager.py`)
   - Specialized recovery mechanisms for different error types
   - Component-specific recovery strategies
   - System resource management and cleanup
   - Recovery statistics tracking

3. **MoonFlowerErrorIntegration** (`utils/error_integration.py`)
   - Integration layer for existing MoonFlower components
   - Component-specific error handling
   - Health monitoring and recommendations
   - Decorator-based error handling

## Error Categories

The system categorizes errors into the following types:

### Network Errors (`ErrorCategory.NETWORK`)
- Connection timeouts
- DNS resolution failures
- SSL certificate issues
- Connection refused errors
- Selenium WebDriver network issues

**Common Recovery Actions:**
- Retry with exponential backoff
- Network adapter reset
- DNS cache flush
- Alternative endpoint testing

### Application Errors (`ErrorCategory.APPLICATION`)
- Process crashes
- Window not found
- Automation failures
- Login failures
- VBS application issues

**Common Recovery Actions:**
- Process restart
- Window recovery
- Application cleanup
- Resource verification

### File System Errors (`ErrorCategory.FILE_SYSTEM`)
- File not found
- Permission denied
- Disk full
- File locked
- Corrupted files

**Common Recovery Actions:**
- File recreation from templates
- Permission fixes
- Disk space cleanup
- File lock resolution
- Backup file restoration

### Data Processing Errors (`ErrorCategory.DATA_PROCESSING`)
- CSV parsing errors
- Excel generation failures
- Data validation errors
- Memory errors
- Encoding issues

**Common Recovery Actions:**
- Alternative parsing methods
- Data chunking
- Memory cleanup
- Skip invalid records
- Encoding fixes

### Email Errors (`ErrorCategory.EMAIL`)
- SMTP connection failures
- Authentication errors
- Attachment issues
- Recipient errors

**Common Recovery Actions:**
- Alternative SMTP servers
- Retry with delays
- Send without attachments
- Recipient validation

### System Errors (`ErrorCategory.SYSTEM`)
- Service failures
- Memory exhaustion
- High CPU usage
- Registry issues
- Configuration errors

**Common Recovery Actions:**
- Service restart
- Resource cleanup
- System health checks
- Configuration validation

## Error Severity Levels

### Low (`ErrorSeverity.LOW`)
- Minor issues with automatic recovery
- Temporary glitches
- Recoverable failures

### Medium (`ErrorSeverity.MEDIUM`)
- Moderate issues requiring intervention
- Component degradation
- Retry-able failures

### High (`ErrorSeverity.HIGH`)
- Serious issues requiring immediate attention
- Component failures
- Service disruptions

### Critical (`ErrorSeverity.CRITICAL`)
- System-threatening issues
- Manual intervention required
- Security or data integrity concerns

## Recovery Actions

### Automatic Recovery (`RecoveryAction.RETRY`)
- Exponential backoff retry logic
- Transient error handling
- Network connectivity restoration

### Component Restart (`RecoveryAction.RESTART_COMPONENT`)
- Individual component restart
- Process cleanup and restart
- Resource reinitialization

### Application Restart (`RecoveryAction.RESTART_APPLICATION`)
- Full application restart
- Process tree termination
- Clean startup sequence

### Service Restart (`RecoveryAction.RESTART_SERVICE`)
- Windows service restart
- Service dependency handling
- Health verification

### System Restart (`RecoveryAction.RESTART_SYSTEM`)
- Full system restart
- Scheduled maintenance restart
- Critical error recovery

### Skip Task (`RecoveryAction.SKIP_TASK`)
- Continue with next task
- Data processing continuation
- Non-critical error handling

### Manual Intervention (`RecoveryAction.MANUAL_INTERVENTION`)
- Human intervention required
- Critical system issues
- Security-related problems

## Usage Examples

### Basic Error Handling

```python
from utils.error_handler import ErrorHandler, ErrorCategory, ErrorSeverity

# Create error handler
handler = ErrorHandler()

# Handle an error
error = handler.handle_error(
    category=ErrorCategory.NETWORK,
    message="Connection timeout to Ruckus controller",
    component="csv_downloader",
    operation="selenium_timeout",
    context={
        "endpoint": "https://51.38.163.73:8443/wsg/",
        "timeout": 30,
        "retry_count": 0
    }
)

print(f"Error handled: {error.id}")
```

### Component-Specific Error Handling

```python
from utils.error_integration import MoonFlowerErrorIntegration

# Create integration instance
integration = MoonFlowerErrorIntegration()

# Handle component error
error = integration.handle_component_error(
    component="vbs_automator",
    message="VBS application crashed during data upload",
    operation="process_crash",
    context={
        "application_name": "AbsonsItERP",
        "process_name": "AbsonsItERP.exe",
        "phase": "data_upload"
    },
    severity=ErrorSeverity.HIGH
)
```

### Decorator-Based Error Handling

```python
from utils.error_integration import csv_downloader_errors, vbs_automator_errors

@csv_downloader_errors("selenium_automation")
def download_csv_files():
    # CSV download logic here
    pass

@vbs_automator_errors("data_upload")
def upload_data_to_vbs():
    # VBS automation logic here
    pass
```

### Error Recovery

```python
from utils.recovery_manager import RecoveryManager
from utils.error_handler import SystemError, ErrorCategory

# Create recovery manager
recovery_manager = RecoveryManager()

# Create error for recovery
error = SystemError(
    category=ErrorCategory.NETWORK,
    severity=ErrorSeverity.MEDIUM,
    message="Network timeout",
    context={"endpoint": "https://example.com"}
)

# Attempt recovery
success = recovery_manager.recover_network_issues(error)
if success:
    error.mark_resolved()
```

## Monitoring and Health Checks

### Error Summary

```python
# Get comprehensive error summary
summary = handler.get_error_summary()

print(f"Total errors: {summary['total_errors']}")
print(f"Recent errors (1h): {summary['recent_errors_1h']}")
print(f"Resolution rate: {summary['resolution_rate_24h']}%")
```

### Component Health Status

```python
# Get component health status
health_status = integration.get_component_health_status()

for component, health in health_status["component_health"].items():
    print(f"{component}: {health['status']} ({health['total_errors']} errors)")

# Get system recommendations
for recommendation in health_status["system_recommendations"]:
    print(f"Recommendation: {recommendation}")
```

### Recovery Statistics

```python
# Get recovery statistics
stats = recovery_manager.get_recovery_stats()

print(f"Recovery attempts: {stats['total_attempts']}")
print(f"Success rate: {stats['success_rate']}%")
```

## Configuration

### Error Handler Configuration

```python
# Configure error handler
handler = ErrorHandler(config_manager=config_manager)

# Set custom thresholds
handler.max_errors_per_hour = 100
handler.max_critical_errors = 10
handler.alert_cooldown_minutes = 30
```

### Recovery Manager Configuration

```python
# Configure recovery manager
recovery_manager = RecoveryManager(config_manager=config_manager)

# Access recovery statistics
stats = recovery_manager.get_recovery_stats()
```

## Integration with Existing Components

### CSV Downloader Integration

```python
from utils.error_integration import handle_csv_downloader_error

try:
    # CSV download logic
    download_csv_files()
except Exception as e:
    handle_csv_downloader_error(
        message="CSV download failed",
        operation="selenium_timeout",
        exception=e,
        context={"url": target_url, "timeout": 30}
    )
```

### VBS Automator Integration

```python
from utils.error_integration import handle_vbs_automator_error

try:
    # VBS automation logic
    automate_vbs_process()
except Exception as e:
    handle_vbs_automator_error(
        message="VBS automation failed",
        operation="process_crash",
        exception=e,
        context={"phase": "data_upload", "window_handle": hwnd}
    )
```

### Excel Generator Integration

```python
from utils.error_integration import handle_excel_generator_error

try:
    # Excel generation logic
    generate_excel_file()
except Exception as e:
    handle_excel_generator_error(
        message="Excel generation failed",
        operation="memory_error",
        exception=e,
        context={"csv_files_count": 8, "memory_usage_mb": 1024}
    )
```

## Logging and Alerting

### Log Files

The error handling system creates detailed log files:

- `EHC_Logs/error_handler_YYYYMMDD.log` - Main error handler logs
- `EHC_Logs/vbs_error_handler.log` - VBS-specific error logs
- `EHC_Logs/recovery_manager.log` - Recovery operation logs

### Log Levels

- **DEBUG**: Detailed execution traces, recovery attempts
- **INFO**: Successful operations, error resolutions
- **WARNING**: Recoverable errors, degraded performance
- **ERROR**: Failed operations, unrecoverable errors
- **CRITICAL**: System failures, manual intervention required

### Email Alerts

Critical and high-severity errors trigger immediate email alerts:

```python
# Configure email alerts
email_settings = {
    'error_recipients': ['admin@moonflower.com', 'support@moonflower.com'],
    'alert_cooldown_minutes': 15
}
```

Alert emails include:
- Error details and context
- System health information
- Recovery attempts and results
- Recommended actions

## Error Export and Analysis

### Export Error Log

```python
# Export error log for analysis
export_file = handler.export_error_log("error_analysis_20240718.json")

# Export includes:
# - Complete error history
# - Error summary statistics
# - System health metrics
# - Recovery success rates
```

### Error Pattern Analysis

The system automatically detects error patterns:

- High error rates (>50 errors/hour)
- Frequent component failures
- Recurring error types
- System degradation trends

## Best Practices

### Error Handling Guidelines

1. **Always provide context**: Include relevant information in error context
2. **Use appropriate severity**: Match severity to actual impact
3. **Enable auto-recovery**: Let the system attempt automatic recovery
4. **Monitor error patterns**: Watch for recurring issues
5. **Review error logs**: Regular log analysis for system health

### Recovery Strategy Guidelines

1. **Implement graceful degradation**: Continue operation with reduced functionality
2. **Use exponential backoff**: Avoid overwhelming failing systems
3. **Validate recovery success**: Ensure recovery actually worked
4. **Log recovery attempts**: Track what works and what doesn't
5. **Escalate when necessary**: Know when to request manual intervention

### Integration Guidelines

1. **Use component-specific handlers**: Leverage specialized error handling
2. **Provide rich context**: Include all relevant error information
3. **Handle exceptions at boundaries**: Catch errors at component interfaces
4. **Test error scenarios**: Verify error handling works correctly
5. **Monitor component health**: Track component-specific error rates

## Testing

### Unit Tests

Run comprehensive error handling tests:

```bash
python tests/test_error_handling.py
```

### Basic Functionality Test

Run basic functionality verification:

```bash
python test_basic_error_handling.py
```

### Integration Tests

Test error handling integration with existing components:

```bash
python tests/test_service_integration.py
```

## Troubleshooting

### Common Issues

1. **High memory usage**: Error tracking consuming too much memory
   - Solution: Clear resolved errors regularly
   - Command: `handler.clear_resolved_errors()`

2. **Alert spam**: Too many error alerts being sent
   - Solution: Increase alert cooldown period
   - Configuration: `handler.alert_cooldown_minutes = 30`

3. **Recovery failures**: Automatic recovery not working
   - Solution: Check recovery strategy implementation
   - Debug: Enable detailed recovery logging

4. **Missing error context**: Errors lack sufficient information
   - Solution: Provide comprehensive context in error calls
   - Example: Include file paths, network endpoints, process IDs

### Debug Mode

Enable debug mode for detailed error tracking:

```python
# Enable debug logging
import logging
logging.getLogger("ErrorHandler").setLevel(logging.DEBUG)
logging.getLogger("RecoveryManager").setLevel(logging.DEBUG)
```

## Performance Considerations

### Memory Management

- Error history is limited to 1000 recent errors
- Resolved errors can be cleared automatically
- Large context objects are truncated in logs

### CPU Usage

- Recovery attempts use exponential backoff
- Alert processing runs in background thread
- Error pattern analysis is throttled

### Disk Usage

- Log files are rotated based on size (50MB limit)
- Old error exports are cleaned up automatically
- Temporary recovery files are removed after use

## Future Enhancements

### Planned Features

1. **Machine Learning Error Prediction**: Predict failures before they occur
2. **Advanced Pattern Recognition**: Detect complex error patterns
3. **Automated Root Cause Analysis**: Identify underlying causes
4. **Integration with External Monitoring**: Connect to monitoring systems
5. **Performance Impact Analysis**: Measure error handling overhead

### Extension Points

The error handling system is designed for extensibility:

- Custom error categories
- Additional recovery strategies
- New alert mechanisms
- Enhanced reporting formats
- Integration with external systems

## Conclusion

The MoonFlower Error Handling and Recovery System provides comprehensive error management capabilities designed for 365-day continuous operation. The system automatically handles common error scenarios, provides detailed logging and monitoring, and enables graceful degradation when issues occur.

For additional support or questions about the error handling system, please refer to the source code documentation or contact the development team.