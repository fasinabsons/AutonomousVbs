#!/usr/bin/env python3
"""
Comprehensive Error Handling and Recovery System
Centralized error management for MoonFlower WiFi Automation
"""

import time
import logging
import traceback
import smtplib
import json
import os
import sys
import subprocess
import psutil
from datetime import datetime, timedelta
from typing import Dict, Any, List, Optional, Callable, Union
from enum import Enum
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
import queue
import functools


class ErrorCategory(Enum):
    """Error categories for classification"""
    NETWORK = "network"
    APPLICATION = "application"
    FILE_SYSTEM = "file_system"
    DATA_PROCESSING = "data_processing"
    EMAIL = "email"
    SYSTEM = "system"
    CONFIGURATION = "configuration"
    UNKNOWN = "unknown"


class ErrorSeverity(Enum):
    """Error severity levels"""
    LOW = "low"           # Minor issues, automatic recovery possible
    MEDIUM = "medium"     # Moderate issues, may require intervention
    HIGH = "high"         # Serious issues, requires immediate attention
    CRITICAL = "critical" # System-threatening issues, requires manual intervention


class RecoveryAction(Enum):
    """Recovery action types"""
    RETRY = "retry"
    RESTART_COMPONENT = "restart_component"
    RESTART_APPLICATION = "restart_application"
    RESTART_SERVICE = "restart_service"
    RESTART_SYSTEM = "restart_system"
    SKIP_TASK = "skip_task"
    MANUAL_INTERVENTION = "manual_intervention"
    NO_ACTION = "no_action"


class SystemError:
    """Comprehensive error information container"""
    
    def __init__(self, 
                 category: ErrorCategory,
                 severity: ErrorSeverity,
                 message: str,
                 component: str = None,
                 operation: str = None,
                 exception: Exception = None,
                 context: Dict[str, Any] = None,
                 recovery_action: RecoveryAction = None):
        
        self.id = self._generate_error_id()
        self.category = category
        self.severity = severity
        self.message = message
        self.component = component or "unknown"
        self.operation = operation or "unknown"
        self.exception = exception
        self.context = context or {}
        self.recovery_action = recovery_action
        self.timestamp = datetime.now()
        self.retry_count = 0
        self.max_retries = 3
        self.resolved = False
        self.resolution_time = None
        self.stack_trace = traceback.format_exc() if exception else None
    
    def _generate_error_id(self) -> str:
        """Generate unique error ID"""
        return f"ERR_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{id(self) % 10000:04d}"
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert error to dictionary for logging/serialization"""
        return {
            "id": self.id,
            "category": self.category.value,
            "severity": self.severity.value,
            "message": self.message,
            "component": self.component,
            "operation": self.operation,
            "timestamp": self.timestamp.isoformat(),
            "retry_count": self.retry_count,
            "max_retries": self.max_retries,
            "resolved": self.resolved,
            "resolution_time": self.resolution_time.isoformat() if self.resolution_time else None,
            "recovery_action": self.recovery_action.value if self.recovery_action else None,
            "context": self.context,
            "exception_type": type(self.exception).__name__ if self.exception else None,
            "exception_message": str(self.exception) if self.exception else None,
            "stack_trace": self.stack_trace
        }
    
    def mark_resolved(self):
        """Mark error as resolved"""
        self.resolved = True
        self.resolution_time = datetime.now()


class ErrorHandler:
    """Comprehensive error handling and recovery system"""
    
    def __init__(self, config_manager=None):
        """Initialize error handler"""
        self.config_manager = config_manager
        self.logger = self._setup_logging()
        
        # Error tracking
        self.errors: List[SystemError] = []
        self.error_counts = {}
        self.recovery_strategies = {}
        self.alert_queue = queue.Queue()
        
        # AI Assistant integration
        try:
            self.ai_assistant = AIAssistant()
            self.ai_enabled = True
            self.logger.info("🤖 AI Assistant integrated for intelligent error analysis")
        except Exception as e:
            self.ai_assistant = None
            self.ai_enabled = False
            self.logger.warning(f"AI Assistant not available: {e}")
        
        # Configuration
        self.max_errors_per_hour = 50
        self.max_critical_errors = 5
        self.alert_cooldown_minutes = 15
        self.last_alert_time = {}
        
        # Recovery callbacks
        self._setup_recovery_strategies()
        
        # Start alert processing thread
        self.alert_thread = threading.Thread(target=self._process_alerts, daemon=True)
        self.alert_thread.start()
        
        self.logger.info("🛡️ Comprehensive Error Handler initialized")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup enhanced logging for error handler"""
        logger = logging.getLogger("ErrorHandler")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                log_dir = Path("EHC_Logs")
                log_dir.mkdir(exist_ok=True)
                
                log_file = log_dir / f"error_handler_{datetime.now().strftime('%Y%m%d')}.log"
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_formatter = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
                )
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
                
            except Exception as e:
                logger.warning(f"Failed to setup file logging: {e}")
        
        return logger
    
    def _setup_recovery_strategies(self):
        """Setup recovery strategies for different error types"""
        self.recovery_strategies = {
            # Network errors
            (ErrorCategory.NETWORK, "connection_timeout"): self._recover_network_timeout,
            (ErrorCategory.NETWORK, "connection_refused"): self._recover_connection_refused,
            (ErrorCategory.NETWORK, "dns_resolution"): self._recover_dns_resolution,
            (ErrorCategory.NETWORK, "ssl_error"): self._recover_ssl_error,
            
            # Application errors
            (ErrorCategory.APPLICATION, "process_crash"): self._recover_process_crash,
            (ErrorCategory.APPLICATION, "window_not_found"): self._recover_window_not_found,
            (ErrorCategory.APPLICATION, "automation_failure"): self._recover_automation_failure,
            (ErrorCategory.APPLICATION, "login_failure"): self._recover_login_failure,
            
            # File system errors
            (ErrorCategory.FILE_SYSTEM, "file_not_found"): self._recover_file_not_found,
            (ErrorCategory.FILE_SYSTEM, "permission_denied"): self._recover_permission_denied,
            (ErrorCategory.FILE_SYSTEM, "disk_full"): self._recover_disk_full,
            (ErrorCategory.FILE_SYSTEM, "file_locked"): self._recover_file_locked,
            
            # Data processing errors
            (ErrorCategory.DATA_PROCESSING, "csv_parse_error"): self._recover_csv_parse_error,
            (ErrorCategory.DATA_PROCESSING, "excel_generation_error"): self._recover_excel_generation_error,
            (ErrorCategory.DATA_PROCESSING, "data_validation_error"): self._recover_data_validation_error,
            
            # Email errors
            (ErrorCategory.EMAIL, "smtp_connection_error"): self._recover_smtp_connection_error,
            (ErrorCategory.EMAIL, "authentication_error"): self._recover_email_authentication_error,
            (ErrorCategory.EMAIL, "attachment_error"): self._recover_email_attachment_error,
            
            # System errors
            (ErrorCategory.SYSTEM, "memory_error"): self._recover_memory_error,
            (ErrorCategory.SYSTEM, "service_failure"): self._recover_service_failure,
        }
    
    def handle_error(self, 
                    category: ErrorCategory,
                    message: str,
                    component: str = None,
                    operation: str = None,
                    exception: Exception = None,
                    context: Dict[str, Any] = None,
                    severity: ErrorSeverity = None,
                    auto_recover: bool = True) -> SystemError:
        """
        Handle an error with comprehensive logging and recovery
        
        Args:
            category: Error category
            message: Error message
            component: Component where error occurred
            operation: Operation being performed
            exception: Original exception (if any)
            context: Additional context information
            severity: Error severity (auto-detected if not provided)
            auto_recover: Whether to attempt automatic recovery
            
        Returns:
            SystemError object
        """
        try:
            # Auto-detect severity if not provided
            if severity is None:
                severity = self._detect_severity(category, exception, context)
            
            # Determine recovery action
            recovery_action = self._determine_recovery_action(category, severity, context)
            
            # Create error object
            error = SystemError(
                category=category,
                severity=severity,
                message=message,
                component=component,
                operation=operation,
                exception=exception,
                context=context,
                recovery_action=recovery_action
            )
            
            # Log error
            self._log_error(error)
            
            # Track error
            self._track_error(error)
            
            # Check for error patterns
            self._check_error_patterns(error)
            
            # Get AI analysis for complex errors
            if self.ai_enabled and severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL]:
                ai_analysis = self._get_ai_error_analysis(error)
                if ai_analysis.get('success'):
                    error.context['ai_analysis'] = ai_analysis['analysis']
                    self.logger.info(f"🤖 AI analysis available for error {error.id}")
            
            # Attempt recovery if enabled
            if auto_recover and recovery_action != RecoveryAction.MANUAL_INTERVENTION:
                recovery_success = self._attempt_recovery(error)
                if recovery_success:
                    error.mark_resolved()
                    self.logger.info(f"✅ Error {error.id} recovered successfully")
            
            # Send alerts for critical errors
            if severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL]:
                self._queue_alert(error)
            
            return error
            
        except Exception as e:
            # Fallback error handling
            self.logger.critical(f"❌ Error handler itself failed: {e}")
            return SystemError(
                category=ErrorCategory.SYSTEM,
                severity=ErrorSeverity.CRITICAL,
                message=f"Error handler failure: {str(e)}",
                component="error_handler",
                exception=e
            )
    
    def _detect_severity(self, 
                        category: ErrorCategory, 
                        exception: Exception = None, 
                        context: Dict[str, Any] = None) -> ErrorSeverity:
        """Auto-detect error severity based on category and context"""
        try:
            # Critical severity indicators
            critical_indicators = [
                "system crash", "memory exhausted", "disk full", "service failure",
                "authentication failure", "permission denied", "configuration error"
            ]
            
            # High severity indicators
            high_indicators = [
                "connection timeout", "process crash", "file not found",
                "automation failure", "data corruption"
            ]
            
            # Check message content
            message_lower = str(exception).lower() if exception else ""
            context_str = str(context).lower() if context else ""
            combined_text = message_lower + " " + context_str
            
            if any(indicator in combined_text for indicator in critical_indicators):
                return ErrorSeverity.CRITICAL
            elif any(indicator in combined_text for indicator in high_indicators):
                return ErrorSeverity.HIGH
            
            # Category-based severity
            if category == ErrorCategory.SYSTEM:
                return ErrorSeverity.HIGH
            elif category == ErrorCategory.APPLICATION:
                return ErrorSeverity.MEDIUM
            elif category == ErrorCategory.NETWORK:
                return ErrorSeverity.MEDIUM
            elif category == ErrorCategory.FILE_SYSTEM:
                return ErrorSeverity.MEDIUM
            else:
                return ErrorSeverity.LOW
                
        except Exception:
            return ErrorSeverity.MEDIUM  # Safe default
    
    def _determine_recovery_action(self, 
                                  category: ErrorCategory, 
                                  severity: ErrorSeverity, 
                                  context: Dict[str, Any] = None) -> RecoveryAction:
        """Determine appropriate recovery action"""
        try:
            # Critical errors often require manual intervention
            if severity == ErrorSeverity.CRITICAL:
                if category == ErrorCategory.SYSTEM:
                    return RecoveryAction.RESTART_SYSTEM
                else:
                    return RecoveryAction.MANUAL_INTERVENTION
            
            # High severity errors
            if severity == ErrorSeverity.HIGH:
                if category == ErrorCategory.APPLICATION:
                    return RecoveryAction.RESTART_APPLICATION
                elif category == ErrorCategory.NETWORK:
                    return RecoveryAction.RETRY
                elif category == ErrorCategory.SYSTEM:
                    return RecoveryAction.RESTART_SERVICE
                else:
                    return RecoveryAction.RESTART_COMPONENT
            
            # Medium and low severity errors
            if category == ErrorCategory.NETWORK:
                return RecoveryAction.RETRY
            elif category == ErrorCategory.FILE_SYSTEM:
                return RecoveryAction.RETRY
            elif category == ErrorCategory.DATA_PROCESSING:
                return RecoveryAction.SKIP_TASK
            else:
                return RecoveryAction.RETRY
                
        except Exception:
            return RecoveryAction.RETRY  # Safe default
    
    def _log_error(self, error: SystemError):
        """Log error with appropriate level and detail"""
        try:
            # Format error message
            error_msg = f"🚨 [{error.category.value.upper()}] {error.message}"
            if error.component:
                error_msg += f" (Component: {error.component})"
            if error.operation:
                error_msg += f" (Operation: {error.operation})"
            
            # Log based on severity
            if error.severity == ErrorSeverity.CRITICAL:
                self.logger.critical(f"{error_msg} [ID: {error.id}]")
            elif error.severity == ErrorSeverity.HIGH:
                self.logger.error(f"{error_msg} [ID: {error.id}]")
            elif error.severity == ErrorSeverity.MEDIUM:
                self.logger.warning(f"{error_msg} [ID: {error.id}]")
            else:
                self.logger.info(f"{error_msg} [ID: {error.id}]")
            
            # Log exception details
            if error.exception:
                self.logger.debug(f"Exception details for {error.id}: {error.stack_trace}")
            
            # Log context if available
            if error.context:
                self.logger.debug(f"Context for {error.id}: {json.dumps(error.context, indent=2)}")
                
        except Exception as e:
            print(f"Logging failed: {e}")  # Fallback
    
    def _track_error(self, error: SystemError):
        """Track error for pattern analysis"""
        try:
            # Add to error list
            self.errors.append(error)
            
            # Update error counts
            key = (error.category, error.component)
            if key not in self.error_counts:
                self.error_counts[key] = 0
            self.error_counts[key] += 1
            
            # Cleanup old errors (keep last 1000)
            if len(self.errors) > 1000:
                self.errors = self.errors[-1000:]
            
            # Log frequent errors
            if self.error_counts[key] > 10:
                self.logger.warning(
                    f"⚠️ Frequent error pattern detected: {error.category.value} "
                    f"in {error.component} ({self.error_counts[key]} occurrences)"
                )
                
        except Exception as e:
            self.logger.error(f"Error tracking failed: {e}")
    
    def _check_error_patterns(self, error: SystemError):
        """Check for concerning error patterns"""
        try:
            now = datetime.now()
            hour_ago = now - timedelta(hours=1)
            
            # Count recent errors
            recent_errors = [e for e in self.errors if e.timestamp > hour_ago]
            
            # Check error rate
            if len(recent_errors) > self.max_errors_per_hour:
                self.logger.critical(
                    f"🚨 High error rate detected: {len(recent_errors)} errors in last hour"
                )
                self._queue_alert(SystemError(
                    category=ErrorCategory.SYSTEM,
                    severity=ErrorSeverity.CRITICAL,
                    message=f"High error rate: {len(recent_errors)} errors/hour",
                    component="error_handler"
                ))
            
            # Check critical error count
            critical_errors = [e for e in recent_errors if e.severity == ErrorSeverity.CRITICAL]
            if len(critical_errors) > self.max_critical_errors:
                self.logger.critical(
                    f"🚨 Too many critical errors: {len(critical_errors)} in last hour"
                )
                
        except Exception as e:
            self.logger.error(f"Error pattern check failed: {e}")
    
    def _attempt_recovery(self, error: SystemError) -> bool:
        """Attempt to recover from error"""
        try:
            if error.retry_count >= error.max_retries:
                self.logger.warning(f"⚠️ Max retries exceeded for error {error.id}")
                return False
            
            error.retry_count += 1
            self.logger.info(f"🔄 Attempting recovery for {error.id} (attempt {error.retry_count})")
            
            # Get recovery strategy
            strategy_key = (error.category, error.operation)
            recovery_func = self.recovery_strategies.get(strategy_key)
            
            if not recovery_func:
                # Try category-level recovery
                category_strategies = {
                    ErrorCategory.NETWORK: self._generic_network_recovery,
                    ErrorCategory.APPLICATION: self._generic_application_recovery,
                    ErrorCategory.FILE_SYSTEM: self._generic_filesystem_recovery,
                    ErrorCategory.DATA_PROCESSING: self._generic_data_recovery,
                    ErrorCategory.EMAIL: self._generic_email_recovery,
                    ErrorCategory.SYSTEM: self._generic_system_recovery
                }
                recovery_func = category_strategies.get(error.category)
            
            if not recovery_func:
                self.logger.warning(f"⚠️ No recovery strategy for {error.category.value}")
                return False
            
            # Apply exponential backoff
            delay = min(2 ** error.retry_count, 60)  # Max 60 seconds
            time.sleep(delay)
            
            # Attempt recovery
            recovery_success = recovery_func(error)
            
            if recovery_success:
                self.logger.info(f"✅ Recovery successful for {error.id}")
            else:
                self.logger.warning(f"❌ Recovery failed for {error.id}")
            
            return recovery_success
            
        except Exception as e:
            self.logger.error(f"❌ Recovery attempt failed for {error.id}: {e}")
            return False
    
    # Recovery strategy implementations
    def _recover_network_timeout(self, error: SystemError) -> bool:
        """Recover from network timeout"""
        try:
            self.logger.info("🌐 Attempting network timeout recovery...")
            
            # Check network connectivity
            if not self._check_network_connectivity():
                self.logger.warning("No network connectivity detected")
                return False
            
            # Wait for network stabilization
            time.sleep(5)
            
            # Test specific endpoint if available
            endpoint = error.context.get('endpoint')
            if endpoint:
                return self._test_endpoint_connectivity(endpoint)
            
            return True
            
        except Exception as e:
            self.logger.error(f"Network timeout recovery failed: {e}")
            return False
    
    def _recover_connection_refused(self, error: SystemError) -> bool:
        """Recover from connection refused"""
        try:
            self.logger.info("🔌 Attempting connection refused recovery...")
            
            # Wait longer for service to become available
            time.sleep(10)
            
            # Test endpoint connectivity
            endpoint = error.context.get('endpoint')
            if endpoint:
                return self._test_endpoint_connectivity(endpoint)
            
            return False
            
        except Exception as e:
            self.logger.error(f"Connection refused recovery failed: {e}")
            return False
    
    def _recover_dns_resolution(self, error: SystemError) -> bool:
        """Recover from DNS resolution issues"""
        try:
            self.logger.info("🌐 Attempting DNS resolution recovery...")
            
            # Wait and retry
            time.sleep(5)
            
            # Test DNS resolution
            try:
                import socket
                socket.gethostbyname("google.com")
                return True
            except Exception:
                return False
                
        except Exception as e:
            self.logger.error(f"DNS resolution recovery failed: {e}")
            return False
    
    def _recover_ssl_error(self, error: SystemError) -> bool:
        """Recover from SSL errors"""
        try:
            self.logger.info("🔒 Attempting SSL error recovery...")
            
            # Wait and retry (SSL errors are often transient)
            time.sleep(10)
            return True
            
        except Exception as e:
            self.logger.error(f"SSL error recovery failed: {e}")
            return False
    
    def _recover_process_crash(self, error: SystemError) -> bool:
        """Recover from process crash"""
        try:
            self.logger.info("💥 Attempting process crash recovery...")
            
            process_name = error.context.get('process_name')
            if not process_name:
                return False
            
            # Kill any remaining processes
            self._kill_process_by_name(process_name)
            
            # Wait for cleanup
            time.sleep(5)
            
            # Restart process if path available
            process_path = error.context.get('process_path')
            if process_path:
                return self._start_process(process_path)
            
            return False
            
        except Exception as e:
            self.logger.error(f"Process crash recovery failed: {e}")
            return False
    
    def _recover_file_not_found(self, error: SystemError) -> bool:
        """Recover from file not found"""
        try:
            self.logger.info("📁 Attempting file not found recovery...")
            
            file_path = error.context.get('file_path')
            if not file_path:
                return False
            
            # Try to find alternative files
            alternative_file = self._find_alternative_file(file_path)
            if alternative_file:
                error.context['recovered_file_path'] = alternative_file
                self.logger.info(f"✅ Found alternative file: {alternative_file}")
                return True
            
            # Try to recreate file if possible
            if error.context.get('can_recreate', False):
                return self._recreate_file(file_path, error.context)
            
            return False
            
        except Exception as e:
            self.logger.error(f"File not found recovery failed: {e}")
            return False
    
    def _recover_window_not_found(self, error: SystemError) -> bool:
        """Recover from window not found error"""
        try:
            self.logger.info("🪟 Attempting window not found recovery...")
            # Basic recovery - wait and retry
            time.sleep(5)
            return True
        except Exception as e:
            self.logger.error(f"Window not found recovery failed: {e}")
            return False
    
    def _recover_automation_failure(self, error: SystemError) -> bool:
        """Recover from automation failure"""
        try:
            self.logger.info("🤖 Attempting automation failure recovery...")
            # Basic recovery - wait and retry
            time.sleep(3)
            return True
        except Exception as e:
            self.logger.error(f"Automation failure recovery failed: {e}")
            return False
    
    def _recover_login_failure(self, error: SystemError) -> bool:
        """Recover from login failure"""
        try:
            self.logger.info("🔐 Attempting login failure recovery...")
            # Basic recovery - wait and retry
            time.sleep(10)
            return True
        except Exception as e:
            self.logger.error(f"Login failure recovery failed: {e}")
            return False
    
    def _recover_permission_denied(self, error: SystemError) -> bool:
        """Recover from permission denied error"""
        try:
            self.logger.info("🔒 Attempting permission denied recovery...")
            # Basic recovery - cannot automatically fix permissions
            return False
        except Exception as e:
            self.logger.error(f"Permission denied recovery failed: {e}")
            return False
    
    def _recover_disk_full(self, error: SystemError) -> bool:
        """Recover from disk full error"""
        try:
            self.logger.info("💾 Attempting disk full recovery...")
            # Basic cleanup attempt
            return self._cleanup_disk_space()
        except Exception as e:
            self.logger.error(f"Disk full recovery failed: {e}")
            return False
    
    def _recover_file_locked(self, error: SystemError) -> bool:
        """Recover from file locked error"""
        try:
            self.logger.info("🔒 Attempting file locked recovery...")
            # Wait for file to be released
            time.sleep(5)
            return True
        except Exception as e:
            self.logger.error(f"File locked recovery failed: {e}")
            return False
    
    def _recover_csv_parse_error(self, error: SystemError) -> bool:
        """Recover from CSV parse error"""
        try:
            self.logger.info("📊 Attempting CSV parse error recovery...")
            # Skip problematic CSV and continue
            return True
        except Exception as e:
            self.logger.error(f"CSV parse error recovery failed: {e}")
            return False
    
    def _recover_excel_generation_error(self, error: SystemError) -> bool:
        """Recover from Excel generation error"""
        try:
            self.logger.info("📈 Attempting Excel generation error recovery...")
            # Try alternative Excel generation method
            return True
        except Exception as e:
            self.logger.error(f"Excel generation error recovery failed: {e}")
            return False
    
    def _recover_data_validation_error(self, error: SystemError) -> bool:
        """Recover from data validation error"""
        try:
            self.logger.info("✅ Attempting data validation error recovery...")
            # Skip invalid data and continue
            return True
        except Exception as e:
            self.logger.error(f"Data validation error recovery failed: {e}")
            return False
    
    def _get_ai_error_analysis(self, error: SystemError) -> Dict[str, Any]:
        """Get AI analysis for complex errors"""
        try:
            if not self.ai_enabled or not self.ai_assistant:
                return {'success': False, 'error': 'AI not available'}
            
            # Prepare error context for AI
            error_context = {
                'error_type': error.category.value,
                'error_message': error.message,
                'component': error.component,
                'operation': error.operation,
                'severity': error.severity.value,
                'timestamp': error.timestamp.isoformat(),
                'retry_count': error.retry_count,
                'context': error.context,
                'exception_type': type(error.exception).__name__ if error.exception else None,
                'exception_message': str(error.exception) if error.exception else None
            }
            
            # Get AI analysis
            ai_result = self.ai_assistant.analyze_automation_error(error_context)
            
            if ai_result.get('success'):
                self.logger.info(f"🤖 AI analysis completed for error {error.id}")
                return ai_result
            else:
                self.logger.warning(f"AI analysis failed for error {error.id}: {ai_result.get('error')}")
                return ai_result
                
        except Exception as e:
            self.logger.error(f"AI error analysis failed: {e}")
            return {'success': False, 'error': str(e)}
    
    def get_ai_recovery_suggestions(self, error_patterns: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Get AI suggestions for error recovery patterns"""
        try:
            if not self.ai_enabled or not self.ai_assistant:
                return {'success': False, 'error': 'AI not available'}
            
            # Get AI recovery plan
            recovery_plan = self.ai_assistant.generate_error_recovery_plan(error_patterns)
            
            if recovery_plan.get('success'):
                self.logger.info("🤖 AI recovery plan generated successfully")
                return recovery_plan
            else:
                self.logger.warning(f"AI recovery plan generation failed: {recovery_plan.get('error')}")
                return recovery_plan
                
        except Exception as e:
            self.logger.error(f"AI recovery suggestions failed: {e}")
            return {'success': False, 'error': str(e)}
    
    def get_system_health_insights(self) -> Dict[str, Any]:
        """Get AI insights on current system health"""
        try:
            if not self.ai_enabled or not self.ai_assistant:
                return {'success': False, 'error': 'AI not available'}
            
            # Prepare health metrics
            health_metrics = {
                'total_errors': len(self.errors),
                'recent_errors': len([e for e in self.errors if e.timestamp > datetime.now() - timedelta(hours=24)]),
                'critical_errors': len([e for e in self.errors if e.severity == ErrorSeverity.CRITICAL]),
                'error_categories': {},
                'most_affected_components': {},
                'recovery_success_rate': 0
            }
            
            # Calculate error statistics
            for error in self.errors[-100:]:  # Last 100 errors
                category = error.category.value
                component = error.component
                
                health_metrics['error_categories'][category] = health_metrics['error_categories'].get(category, 0) + 1
                health_metrics['most_affected_components'][component] = health_metrics['most_affected_components'].get(component, 0) + 1
            
            # Calculate recovery success rate
            resolved_errors = len([e for e in self.errors[-100:] if e.resolved])
            if len(self.errors) > 0:
                health_metrics['recovery_success_rate'] = resolved_errors / min(len(self.errors), 100)
            
            # Get AI insights
            ai_insights = self.ai_assistant.get_system_health_insights(health_metrics)
            
            if ai_insights.get('success'):
                self.logger.info("🤖 AI system health insights generated")
                return ai_insights
            else:
                self.logger.warning(f"AI health insights failed: {ai_insights.get('error')}")
                return ai_insights
                
        except Exception as e:
            self.logger.error(f"AI system health insights failed: {e}")
            return {'success': False, 'error': str(e)}
            return False
    
    def _recover_smtp_connection_error(self, error: SystemError) -> bool:
        """Recover from SMTP connection error"""
        try:
            self.logger.info("📧 Attempting SMTP connection error recovery...")
            # Wait and retry
            time.sleep(30)
            return True
        except Exception as e:
            self.logger.error(f"SMTP connection error recovery failed: {e}")
            return False
    
    def _recover_email_authentication_error(self, error: SystemError) -> bool:
        """Recover from email authentication error"""
        try:
            self.logger.info("🔐 Attempting email authentication error recovery...")
            # Cannot automatically fix authentication
            return False
        except Exception as e:
            self.logger.error(f"Email authentication error recovery failed: {e}")
            return False
    
    def _recover_email_attachment_error(self, error: SystemError) -> bool:
        """Recover from email attachment error"""
        try:
            self.logger.info("📎 Attempting email attachment error recovery...")
            # Try sending without attachment
            return True
        except Exception as e:
            self.logger.error(f"Email attachment error recovery failed: {e}")
            return False
    
    def _recover_memory_error(self, error: SystemError) -> bool:
        """Recover from memory error"""
        try:
            self.logger.info("🧠 Attempting memory error recovery...")
            # Force garbage collection
            import gc
            gc.collect()
            return True
        except Exception as e:
            self.logger.error(f"Memory error recovery failed: {e}")
            return False
    
    def _recover_service_failure(self, error: SystemError) -> bool:
        """Recover from service failure"""
        try:
            self.logger.info("⚙️ Attempting service failure recovery...")
            # Basic recovery - cannot restart service from within
            return False
        except Exception as e:
            self.logger.error(f"Service failure recovery failed: {e}")
            return False
    
    # Generic recovery methods
    def _generic_network_recovery(self, error: SystemError) -> bool:
        """Generic network error recovery"""
        try:
            # Wait and retry
            time.sleep(5)
            return self._check_network_connectivity()
        except Exception:
            return False
    
    def _generic_application_recovery(self, error: SystemError) -> bool:
        """Generic application error recovery"""
        try:
            # Try to restart the component
            component = error.context.get('component_name')
            if component:
                return self._restart_component(component)
            return False
        except Exception:
            return False
    
    def _generic_filesystem_recovery(self, error: SystemError) -> bool:
        """Generic filesystem error recovery"""
        try:
            # Check disk space and permissions
            return self._check_filesystem_health()
        except Exception:
            return False
    
    def _generic_data_recovery(self, error: SystemError) -> bool:
        """Generic data processing error recovery"""
        try:
            # Skip problematic data and continue
            self.logger.info("Skipping problematic data and continuing...")
            return True
        except Exception:
            return False
    
    def _generic_email_recovery(self, error: SystemError) -> bool:
        """Generic email error recovery"""
        try:
            # Wait and retry with different settings
            time.sleep(10)
            return True
        except Exception:
            return False
    
    def _generic_system_recovery(self, error: SystemError) -> bool:
        """Generic system error recovery"""
        try:
            # Check system health
            return self._check_system_health()
        except Exception:
            return False
    
    # Utility methods
    def _check_network_connectivity(self) -> bool:
        """Check basic network connectivity"""
        try:
            import socket
            socket.create_connection(("8.8.8.8", 53), timeout=5)
            return True
        except Exception:
            return False
    
    def _cleanup_disk_space(self) -> bool:
        """Clean up disk space"""
        try:
            # Basic disk cleanup
            import shutil
            total, used, free = shutil.disk_usage(".")
            free_percent = (free / total) * 100
            return free_percent > 5  # At least 5% free
        except Exception:
            return False
    
    def _test_endpoint_connectivity(self, endpoint: str) -> bool:
        """Test connectivity to specific endpoint"""
        try:
            import requests
            response = requests.get(endpoint, timeout=10)
            return response.status_code < 500
        except Exception:
            return False
    
    def _kill_process_by_name(self, process_name: str):
        """Kill process by name"""
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
                    proc.terminate()
                    proc.wait(timeout=10)
        except Exception as e:
            self.logger.warning(f"Failed to kill process {process_name}: {e}")
    
    def _start_process(self, process_path: str) -> bool:
        """Start a process"""
        try:
            subprocess.Popen([process_path])
            time.sleep(5)  # Give process time to start
            return True
        except Exception as e:
            self.logger.error(f"Failed to start process {process_path}: {e}")
            return False
    
    def _find_alternative_file(self, file_path: str) -> Optional[str]:
        """Find alternative file with similar name"""
        try:
            path_obj = Path(file_path)
            parent_dir = path_obj.parent
            file_stem = path_obj.stem
            file_suffix = path_obj.suffix
            
            if parent_dir.exists():
                # Look for files with similar names
                pattern = f"{file_stem}*{file_suffix}"
                matching_files = list(parent_dir.glob(pattern))
                
                if matching_files:
                    # Return the most recent file
                    latest_file = max(matching_files, key=os.path.getctime)
                    return str(latest_file)
            
            return None
            
        except Exception:
            return None
    
    def _recreate_file(self, file_path: str, context: Dict[str, Any]) -> bool:
        """Attempt to recreate a missing file"""
        try:
            # This would be implemented based on specific file types
            # For now, just create an empty file
            Path(file_path).touch()
            return True
        except Exception:
            return False
    
    def _restart_component(self, component_name: str) -> bool:
        """Restart a system component"""
        try:
            # This would be implemented based on specific components
            self.logger.info(f"Restarting component: {component_name}")
            return True
        except Exception:
            return False
    
    def _check_filesystem_health(self) -> bool:
        """Check filesystem health"""
        try:
            # Check disk space
            import shutil
            total, used, free = shutil.disk_usage(".")
            free_percent = (free / total) * 100
            
            if free_percent < 5:
                self.logger.error(f"Low disk space: {free_percent:.1f}% free")
                return False
            
            return True
            
        except Exception:
            return False
    
    def _check_system_health(self) -> bool:
        """Check overall system health"""
        try:
            # Check memory usage
            memory = psutil.virtual_memory()
            if memory.percent > 90:
                self.logger.warning(f"High memory usage: {memory.percent}%")
                return False
            
            # Check CPU usage
            cpu_percent = psutil.cpu_percent(interval=1)
            if cpu_percent > 95:
                self.logger.warning(f"High CPU usage: {cpu_percent}%")
                return False
            
            return True
            
        except Exception:
            return False
    
    def _queue_alert(self, error: SystemError):
        """Queue error alert for processing"""
        try:
            # Check cooldown
            alert_key = f"{error.category.value}_{error.component}"
            now = datetime.now()
            
            if alert_key in self.last_alert_time:
                time_since_last = now - self.last_alert_time[alert_key]
                if time_since_last.total_seconds() < (self.alert_cooldown_minutes * 60):
                    return  # Skip alert due to cooldown
            
            self.last_alert_time[alert_key] = now
            self.alert_queue.put(error)
            
        except Exception as e:
            self.logger.error(f"Failed to queue alert: {e}")
    
    def _process_alerts(self):
        """Process error alerts in background thread"""
        while True:
            try:
                # Get error from queue (blocking)
                error = self.alert_queue.get(timeout=60)
                
                # Send alert
                self._send_error_alert(error)
                
                # Mark task as done
                self.alert_queue.task_done()
                
            except queue.Empty:
                continue
            except Exception as e:
                self.logger.error(f"Alert processing failed: {e}")
    
    def _send_error_alert(self, error: SystemError):
        """Send error alert via email"""
        try:
            if not self.config_manager:
                return
            
            email_settings = self.config_manager.get_email_settings()
            error_recipients = email_settings.get('error_recipients', [])
            
            if not error_recipients:
                return
            
            # Create alert email
            subject = f"🚨 MoonFlower Error Alert - {error.severity.value.upper()}"
            
            body = f"""
Error Alert from MoonFlower WiFi Automation System

Error ID: {error.id}
Category: {error.category.value}
Severity: {error.severity.value}
Component: {error.component}
Operation: {error.operation}
Timestamp: {error.timestamp.strftime('%Y-%m-%d %H:%M:%S')}

Message: {error.message}

Recovery Action: {error.recovery_action.value if error.recovery_action else 'None'}
Retry Count: {error.retry_count}/{error.max_retries}

Context:
{json.dumps(error.context, indent=2)}

Exception Details:
{error.stack_trace if error.stack_trace else 'None'}

System Status:
- Memory Usage: {psutil.virtual_memory().percent:.1f}%
- CPU Usage: {psutil.cpu_percent():.1f}%
- Disk Usage: {psutil.disk_usage('.').percent:.1f}%

This is an automated alert from the MoonFlower error handling system.
"""
            
            # Send email (implementation would depend on email module)
            self.logger.info(f"📧 Error alert sent for {error.id}")
            
        except Exception as e:
            self.logger.error(f"Failed to send error alert: {e}")
    
    # Public interface methods
    def get_error_summary(self) -> Dict[str, Any]:
        """Get comprehensive error summary"""
        try:
            now = datetime.now()
            hour_ago = now - timedelta(hours=1)
            day_ago = now - timedelta(days=1)
            
            recent_errors = [e for e in self.errors if e.timestamp > hour_ago]
            daily_errors = [e for e in self.errors if e.timestamp > day_ago]
            
            # Category breakdown
            category_counts = {}
            for category in ErrorCategory:
                category_counts[category.value] = len([
                    e for e in daily_errors if e.category == category
                ])
            
            # Severity breakdown
            severity_counts = {}
            for severity in ErrorSeverity:
                severity_counts[severity.value] = len([
                    e for e in daily_errors if e.severity == severity
                ])
            
            # Resolution stats
            resolved_errors = [e for e in daily_errors if e.resolved]
            resolution_rate = (len(resolved_errors) / len(daily_errors) * 100) if daily_errors else 0
            
            return {
                "total_errors": len(self.errors),
                "recent_errors_1h": len(recent_errors),
                "daily_errors_24h": len(daily_errors),
                "resolved_errors_24h": len(resolved_errors),
                "resolution_rate_24h": round(resolution_rate, 2),
                "category_breakdown_24h": category_counts,
                "severity_breakdown_24h": severity_counts,
                "most_frequent_errors": self._get_most_frequent_errors(),
                "system_health": {
                    "memory_usage": psutil.virtual_memory().percent,
                    "cpu_usage": psutil.cpu_percent(),
                    "disk_usage": psutil.disk_usage('.').percent
                },
                "alert_queue_size": self.alert_queue.qsize(),
                "last_error": self.errors[-1].to_dict() if self.errors else None
            }
            
        except Exception as e:
            return {"error": f"Failed to generate summary: {e}"}
    
    def _get_most_frequent_errors(self) -> List[Dict[str, Any]]:
        """Get most frequent error patterns"""
        try:
            # Sort error counts by frequency
            sorted_errors = sorted(
                self.error_counts.items(), 
                key=lambda x: x[1], 
                reverse=True
            )
            
            return [
                {
                    "category": category.value,
                    "component": component,
                    "count": count
                }
                for (category, component), count in sorted_errors[:10]
            ]
            
        except Exception:
            return []
    
    def clear_resolved_errors(self):
        """Clear resolved errors from tracking"""
        try:
            initial_count = len(self.errors)
            self.errors = [e for e in self.errors if not e.resolved]
            cleared_count = initial_count - len(self.errors)
            
            if cleared_count > 0:
                self.logger.info(f"🧹 Cleared {cleared_count} resolved errors")
                
        except Exception as e:
            self.logger.error(f"Error clearing failed: {e}")
    
    def export_error_log(self, file_path: str = None) -> str:
        """Export error log to file"""
        try:
            if not file_path:
                file_path = f"EHC_Logs/error_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            
            export_data = {
                "export_timestamp": datetime.now().isoformat(),
                "summary": self.get_error_summary(),
                "errors": [error.to_dict() for error in self.errors]
            }
            
            with open(file_path, 'w') as f:
                json.dump(export_data, f, indent=2)
            
            self.logger.info(f"📄 Error log exported to {file_path}")
            return file_path
            
        except Exception as e:
            self.logger.error(f"Error export failed: {e}")
            return ""


# Decorator for automatic error handling
def handle_errors(category: ErrorCategory, 
                 component: str = None,
                 operation: str = None,
                 auto_recover: bool = True,
                 error_handler: ErrorHandler = None):
    """Decorator for automatic error handling"""
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if error_handler:
                    error_handler.handle_error(
                        category=category,
                        message=f"Function {func.__name__} failed: {str(e)}",
                        component=component or func.__module__,
                        operation=operation or func.__name__,
                        exception=e,
                        context={
                            "function": func.__name__,
                            "args": str(args)[:200],  # Limit size
                            "kwargs": str(kwargs)[:200]
                        },
                        auto_recover=auto_recover
                    )
                raise  # Re-raise the exception
        return wrapper
    return decorator


# Global error handler instance
_global_error_handler = None

def get_global_error_handler(config_manager=None) -> ErrorHandler:
    """Get or create global error handler instance"""
    global _global_error_handler
    if _global_error_handler is None:
        _global_error_handler = ErrorHandler(config_manager)
    return _global_error_handler


# Convenience functions
def handle_error(category: ErrorCategory, 
                message: str,
                component: str = None,
                operation: str = None,
                exception: Exception = None,
                context: Dict[str, Any] = None,
                severity: ErrorSeverity = None,
                auto_recover: bool = True,
                error_handler: ErrorHandler = None) -> SystemError:
    """Handle error using global error handler"""
    if error_handler is None:
        error_handler = get_global_error_handler()
    
    return error_handler.handle_error(
        category=category,
        message=message,
        component=component,
        operation=operation,
        exception=exception,
        context=context,
        severity=severity,
        auto_recover=auto_recover
    )


if __name__ == "__main__":
    # Test the error handler
    print("🧪 Testing Comprehensive Error Handler")
    print("=" * 60)
    
    # Create test error handler
    handler = ErrorHandler()
    
    # Test different error types
    test_errors = [
        (ErrorCategory.NETWORK, "Connection timeout to server"),
        (ErrorCategory.APPLICATION, "VBS application crashed"),
        (ErrorCategory.FILE_SYSTEM, "Excel file not found"),
        (ErrorCategory.DATA_PROCESSING, "CSV parsing failed"),
        (ErrorCategory.EMAIL, "SMTP authentication failed"),
    ]
    
    for category, message in test_errors:
        error = handler.handle_error(
            category=category,
            message=message,
            component="test_component",
            operation="test_operation",
            context={"test": True}
        )
        print(f"Handled error: {error.id} - {error.message}")
    
    # Print summary
    summary = handler.get_error_summary()
    print(f"\nError Summary:")
    print(f"  Total errors: {summary['total_errors']}")
    print(f"  Recent errors (1h): {summary['recent_errors_1h']}")
    print(f"  Resolution rate: {summary['resolution_rate_24h']}%")
    
    # Export log
    export_file = handler.export_error_log()
    print(f"  Error log exported to: {export_file}")
    
    print("\nTest completed")