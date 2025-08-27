#!/usr/bin/env python3
"""
Error Handling and Recovery Service for VBS Computer Vision Automation
Provides comprehensive error management, automatic recovery, and diagnostic capabilities
"""

import time
import logging
import os
import json
import traceback
from typing import List, Dict, Tuple, Optional, Any, Callable
from dataclasses import dataclass, asdict
from enum import Enum
import cv2
import numpy as np
import win32gui
from .config_loader import get_cv_config

class ErrorCategory(Enum):
    """Categories of automation errors"""
    ELEMENT_NOT_FOUND = "element_not_found"
    LOW_CONFIDENCE = "low_confidence"
    SCREENSHOT_FAILURE = "screenshot_failure"
    TIMEOUT = "timeout"
    COORDINATE_ERROR = "coordinate_error"
    OCR_ERROR = "ocr_error"
    TEMPLATE_ERROR = "template_error"
    SYSTEM_ERROR = "system_error"
    CONFIGURATION_ERROR = "configuration_error"
    NETWORK_ERROR = "network_error"
    UNKNOWN = "unknown"

class RecoveryStrategy(Enum):
    """Available recovery strategies"""
    RETRY_SAME_METHOD = "retry_same_method"
    TRY_NEXT_METHOD = "try_next_method"
    ADJUST_PARAMETERS = "adjust_parameters"
    CAPTURE_NEW_TEMPLATE = "capture_new_template"
    FALLBACK_TO_COORDINATES = "fallback_to_coordinates"
    WAIT_AND_RETRY = "wait_and_retry"
    RESTART_SERVICE = "restart_service"
    MANUAL_INTERVENTION = "manual_intervention"
    ABORT_OPERATION = "abort_operation"

@dataclass
class ErrorContext:
    """Context information for an error"""
    timestamp: float
    error_category: ErrorCategory
    error_message: str
    method_used: str
    action_type: str
    target_element: Optional[str] = None
    confidence_score: Optional[float] = None
    screenshot_path: Optional[str] = None
    stack_trace: Optional[str] = None
    system_info: Optional[Dict[str, Any]] = None
    retry_count: int = 0
    previous_attempts: Optional[List[Dict[str, Any]]] = None

@dataclass
class RecoveryAction:
    """Represents a recovery action to be taken"""
    strategy: RecoveryStrategy
    parameters: Dict[str, Any]
    priority: int
    description: str
    estimated_success_rate: float

@dataclass
class RecoveryResult:
    """Result of a recovery attempt"""
    success: bool
    strategy_used: RecoveryStrategy
    execution_time: float
    new_parameters: Optional[Dict[str, Any]] = None
    error_message: Optional[str] = None
    should_retry: bool = True

class CVErrorHandler:
    """Comprehensive error handling and recovery service for CV automation"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        
        # Configuration
        self.max_recovery_attempts = self.config.get('smart_automation.max_retries', 3)
        self.screenshot_on_error = self.config.get('smart_automation.screenshot_on_error', True)
        self.auto_parameter_adjustment = self.config.get('debugging.auto_parameter_adjustment', True)
        
        # Error tracking
        self.error_history = []
        self.recovery_stats = {
            'total_errors': 0,
            'successful_recoveries': 0,
            'recovery_strategy_success': {strategy.value: 0 for strategy in RecoveryStrategy},
            'error_category_counts': {category.value: 0 for category in ErrorCategory},
            'average_recovery_time': 0.0
        }
        
        # Parameter adjustment history
        self.parameter_adjustments = {}
        
        # Recovery strategy definitions
        self.recovery_strategies = self._initialize_recovery_strategies()
        
        self.logger.info("CV Error Handler initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for error handler"""
        logger = logging.getLogger("CVErrorHandler")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler for error logs
            try:
                log_file = "EHC_Logs/cv_error_handler.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def handle_error(self, error: Exception, context: Dict[str, Any]) -> RecoveryResult:
        """Main error handling entry point"""
        start_time = time.time()
        self.recovery_stats['total_errors'] += 1
        
        try:
            # Create error context
            error_context = self._create_error_context(error, context)
            
            # Log error details
            self._log_error(error_context)
            
            # Capture diagnostic information
            self._capture_diagnostics(error_context)
            
            # Determine recovery strategy
            recovery_actions = self._determine_recovery_strategy(error_context)
            
            # Execute recovery actions
            for action in recovery_actions:
                recovery_result = self._execute_recovery_action(action, error_context)
                
                recovery_result.execution_time = time.time() - start_time
                
                if recovery_result.success:
                    self.recovery_stats['successful_recoveries'] += 1
                    self.recovery_stats['recovery_strategy_success'][action.strategy.value] += 1
                    self._update_average_recovery_time(recovery_result.execution_time)
                    
                    self.logger.info(f"Recovery successful using strategy: {action.strategy.value}")
                    return recovery_result
                else:
                    self.logger.warning(f"Recovery strategy {action.strategy.value} failed: {recovery_result.error_message}")
            
            # All recovery strategies failed
            return RecoveryResult(
                success=False,
                strategy_used=RecoveryStrategy.ABORT_OPERATION,
                execution_time=time.time() - start_time,
                error_message="All recovery strategies failed",
                should_retry=False
            )
            
        except Exception as recovery_error:
            self.logger.error(f"Error handler itself failed: {recovery_error}")
            return RecoveryResult(
                success=False,
                strategy_used=RecoveryStrategy.ABORT_OPERATION,
                execution_time=time.time() - start_time,
                error_message=f"Error handler failure: {str(recovery_error)}",
                should_retry=False
            )
    
    def _create_error_context(self, error: Exception, context: Dict[str, Any]) -> ErrorContext:
        """Create comprehensive error context"""
        error_category = self._categorize_error(error, context)
        
        return ErrorContext(
            timestamp=time.time(),
            error_category=error_category,
            error_message=str(error),
            method_used=context.get('method_used', 'unknown'),
            action_type=context.get('action_type', 'unknown'),
            target_element=context.get('target_element'),
            confidence_score=context.get('confidence_score'),
            screenshot_path=context.get('screenshot_path'),
            stack_trace=traceback.format_exc(),
            system_info=self._get_system_info(),
            retry_count=context.get('retry_count', 0),
            previous_attempts=context.get('previous_attempts', [])
        )
    
    def _categorize_error(self, error: Exception, context: Dict[str, Any]) -> ErrorCategory:
        """Categorize error for appropriate recovery strategy"""
        error_message = str(error).lower()
        
        if "not found" in error_message or "no matches" in error_message:
            return ErrorCategory.ELEMENT_NOT_FOUND
        elif "confidence" in error_message or "threshold" in error_message:
            return ErrorCategory.LOW_CONFIDENCE
        elif "screenshot" in error_message or "capture" in error_message:
            return ErrorCategory.SCREENSHOT_FAILURE
        elif "timeout" in error_message or "timed out" in error_message:
            return ErrorCategory.TIMEOUT
        elif "coordinate" in error_message or "position" in error_message:
            return ErrorCategory.COORDINATE_ERROR
        elif "ocr" in error_message or "tesseract" in error_message:
            return ErrorCategory.OCR_ERROR
        elif "template" in error_message or "matching" in error_message:
            return ErrorCategory.TEMPLATE_ERROR
        elif "config" in error_message or "setting" in error_message:
            return ErrorCategory.CONFIGURATION_ERROR
        elif "network" in error_message or "connection" in error_message:
            return ErrorCategory.NETWORK_ERROR
        elif isinstance(error, (OSError, SystemError)):
            return ErrorCategory.SYSTEM_ERROR
        else:
            return ErrorCategory.UNKNOWN
    
    def _log_error(self, error_context: ErrorContext):
        """Log error with comprehensive details"""
        self.logger.error(f"CV Automation Error - Category: {error_context.error_category.value}")
        self.logger.error(f"Method: {error_context.method_used}, Action: {error_context.action_type}")
        self.logger.error(f"Target: {error_context.target_element}, Confidence: {error_context.confidence_score}")
        self.logger.error(f"Message: {error_context.error_message}")
        
        if error_context.screenshot_path:
            self.logger.error(f"Screenshot saved: {error_context.screenshot_path}")
        
        # Add to error history
        self.error_history.append(error_context)
        self.recovery_stats['error_category_counts'][error_context.error_category.value] += 1
        
        # Save detailed error report
        self._save_error_report(error_context)
    
    def _capture_diagnostics(self, error_context: ErrorContext):
        """Capture diagnostic information for error analysis"""
        try:
            # Capture screenshot if enabled and not already captured
            if self.screenshot_on_error and not error_context.screenshot_path:
                screenshot_path = self._capture_error_screenshot(error_context)
                error_context.screenshot_path = screenshot_path
            
            # Capture window information
            try:
                foreground_window = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(foreground_window)
                window_rect = win32gui.GetWindowRect(foreground_window)
                
                error_context.system_info = error_context.system_info or {}
                error_context.system_info.update({
                    'active_window': window_title,
                    'window_rect': window_rect,
                    'window_handle': foreground_window
                })
            except Exception as e:
                self.logger.warning(f"Failed to capture window info: {e}")
            
        except Exception as e:
            self.logger.warning(f"Failed to capture diagnostics: {e}")
    
    def _determine_recovery_strategy(self, error_context: ErrorContext) -> List[RecoveryAction]:
        """Determine appropriate recovery strategies based on error context"""
        strategies = []
        
        # Get base strategies for error category
        base_strategies = self.recovery_strategies.get(error_context.error_category, [])
        
        for strategy_info in base_strategies:
            strategy = RecoveryStrategy(strategy_info['strategy'])
            
            # Adjust parameters based on error context
            parameters = strategy_info['parameters'].copy()
            
            if strategy == RecoveryStrategy.ADJUST_PARAMETERS:
                parameters.update(self._get_parameter_adjustments(error_context))
            elif strategy == RecoveryStrategy.WAIT_AND_RETRY:
                # Increase wait time based on retry count
                parameters['wait_time'] = parameters.get('wait_time', 1.0) * (error_context.retry_count + 1)
            
            # Calculate success rate based on history
            success_rate = self._calculate_strategy_success_rate(strategy, error_context)
            
            action = RecoveryAction(
                strategy=strategy,
                parameters=parameters,
                priority=strategy_info['priority'],
                description=strategy_info['description'],
                estimated_success_rate=success_rate
            )
            
            strategies.append(action)
        
        # Sort by priority and success rate
        strategies.sort(key=lambda x: (x.priority, -x.estimated_success_rate))
        
        return strategies
    
    def _execute_recovery_action(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Execute a specific recovery action"""
        start_time = time.time()
        
        try:
            self.logger.info(f"Executing recovery strategy: {action.strategy.value}")
            
            if action.strategy == RecoveryStrategy.RETRY_SAME_METHOD:
                return self._retry_same_method(action, error_context)
            
            elif action.strategy == RecoveryStrategy.TRY_NEXT_METHOD:
                return self._try_next_method(action, error_context)
            
            elif action.strategy == RecoveryStrategy.ADJUST_PARAMETERS:
                return self._adjust_parameters(action, error_context)
            
            elif action.strategy == RecoveryStrategy.CAPTURE_NEW_TEMPLATE:
                return self._capture_new_template(action, error_context)
            
            elif action.strategy == RecoveryStrategy.FALLBACK_TO_COORDINATES:
                return self._fallback_to_coordinates(action, error_context)
            
            elif action.strategy == RecoveryStrategy.WAIT_AND_RETRY:
                return self._wait_and_retry(action, error_context)
            
            elif action.strategy == RecoveryStrategy.RESTART_SERVICE:
                return self._restart_service(action, error_context)
            
            elif action.strategy == RecoveryStrategy.MANUAL_INTERVENTION:
                return self._request_manual_intervention(action, error_context)
            
            else:
                return RecoveryResult(
                    success=False,
                    strategy_used=action.strategy,
                    execution_time=time.time() - start_time,
                    error_message=f"Unknown recovery strategy: {action.strategy.value}",
                    should_retry=False
                )
                
        except Exception as e:
            return RecoveryResult(
                success=False,
                strategy_used=action.strategy,
                execution_time=time.time() - start_time,
                error_message=f"Recovery action failed: {str(e)}",
                should_retry=True
            )
    
    def _retry_same_method(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Retry the same method with slight delay"""
        wait_time = action.parameters.get('wait_time', 1.0)
        time.sleep(wait_time)
        
        return RecoveryResult(
            success=True,
            strategy_used=action.strategy,
            execution_time=wait_time,
            should_retry=True
        )
    
    def _try_next_method(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Signal to try the next available method"""
        return RecoveryResult(
            success=True,
            strategy_used=action.strategy,
            execution_time=0.1,
            new_parameters={'try_next_method': True},
            should_retry=True
        )
    
    def _adjust_parameters(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Adjust detection parameters for better results"""
        adjustments = action.parameters.get('adjustments', {})
        
        # Apply parameter adjustments
        new_parameters = {}
        
        if error_context.error_category == ErrorCategory.LOW_CONFIDENCE:
            # Lower confidence threshold
            current_threshold = error_context.confidence_score or 0.8
            new_threshold = max(0.5, current_threshold - 0.1)
            new_parameters['confidence_threshold'] = new_threshold
            
        elif error_context.error_category == ErrorCategory.OCR_ERROR:
            # Adjust OCR parameters
            new_parameters.update({
                'ocr_preprocessing': True,
                'gaussian_blur_kernel': 5,
                'adaptive_threshold_block_size': 15
            })
            
        elif error_context.error_category == ErrorCategory.TEMPLATE_ERROR:
            # Adjust template matching parameters
            new_parameters.update({
                'template_scale_factors': [0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3],
                'template_threshold': 0.6
            })
        
        # Store adjustment for future reference
        element_key = error_context.target_element or 'unknown'
        if element_key not in self.parameter_adjustments:
            self.parameter_adjustments[element_key] = []
        
        self.parameter_adjustments[element_key].append({
            'timestamp': time.time(),
            'adjustments': new_parameters,
            'error_category': error_context.error_category.value
        })
        
        return RecoveryResult(
            success=True,
            strategy_used=action.strategy,
            execution_time=0.1,
            new_parameters=new_parameters,
            should_retry=True
        )
    
    def _capture_new_template(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Attempt to capture a new template for the failed element"""
        # This would typically involve user interaction or automated template capture
        # For now, we'll simulate the process
        
        self.logger.info(f"Template capture requested for element: {error_context.target_element}")
        
        # In a real implementation, this would:
        # 1. Capture current screen
        # 2. Allow user to select element region
        # 3. Save new template
        # 4. Update template database
        
        return RecoveryResult(
            success=False,  # Requires manual intervention
            strategy_used=action.strategy,
            execution_time=0.1,
            error_message="Template capture requires manual intervention",
            should_retry=False
        )
    
    def _fallback_to_coordinates(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Fallback to coordinate-based automation"""
        fallback_coordinates = action.parameters.get('coordinates')
        
        if fallback_coordinates:
            return RecoveryResult(
                success=True,
                strategy_used=action.strategy,
                execution_time=0.1,
                new_parameters={'fallback_coordinates': fallback_coordinates},
                should_retry=True
            )
        else:
            return RecoveryResult(
                success=False,
                strategy_used=action.strategy,
                execution_time=0.1,
                error_message="No fallback coordinates available",
                should_retry=False
            )
    
    def _wait_and_retry(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Wait for a specified time and retry"""
        wait_time = action.parameters.get('wait_time', 2.0)
        
        self.logger.info(f"Waiting {wait_time} seconds before retry")
        time.sleep(wait_time)
        
        return RecoveryResult(
            success=True,
            strategy_used=action.strategy,
            execution_time=wait_time,
            should_retry=True
        )
    
    def _restart_service(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Restart the automation service"""
        # This would typically restart the CV services
        self.logger.warning("Service restart requested - not implemented in this context")
        
        return RecoveryResult(
            success=False,
            strategy_used=action.strategy,
            execution_time=0.1,
            error_message="Service restart not implemented",
            should_retry=False
        )
    
    def _request_manual_intervention(self, action: RecoveryAction, error_context: ErrorContext) -> RecoveryResult:
        """Request manual intervention"""
        self.logger.error(f"Manual intervention required for: {error_context.error_message}")
        
        # Save detailed error report for manual review
        self._save_manual_intervention_report(error_context)
        
        return RecoveryResult(
            success=False,
            strategy_used=action.strategy,
            execution_time=0.1,
            error_message="Manual intervention required",
            should_retry=False
        )    
def _initialize_recovery_strategies(self) -> Dict[ErrorCategory, List[Dict[str, Any]]]:
        """Initialize recovery strategies for each error category"""
        return {
            ErrorCategory.ELEMENT_NOT_FOUND: [
                {
                    'strategy': 'adjust_parameters',
                    'priority': 1,
                    'description': 'Lower confidence threshold and adjust detection parameters',
                    'parameters': {'adjustments': {'confidence_threshold': -0.1}}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 2,
                    'description': 'Try alternative detection method',
                    'parameters': {}
                },
                {
                    'strategy': 'wait_and_retry',
                    'priority': 3,
                    'description': 'Wait for UI to stabilize and retry',
                    'parameters': {'wait_time': 2.0}
                },
                {
                    'strategy': 'capture_new_template',
                    'priority': 4,
                    'description': 'Capture new template for element',
                    'parameters': {}
                },
                {
                    'strategy': 'fallback_to_coordinates',
                    'priority': 5,
                    'description': 'Use hardcoded coordinates as fallback',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.LOW_CONFIDENCE: [
                {
                    'strategy': 'adjust_parameters',
                    'priority': 1,
                    'description': 'Lower confidence threshold',
                    'parameters': {'adjustments': {'confidence_threshold': -0.1}}
                },
                {
                    'strategy': 'retry_same_method',
                    'priority': 2,
                    'description': 'Retry with same parameters',
                    'parameters': {'wait_time': 0.5}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 3,
                    'description': 'Try alternative detection method',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.SCREENSHOT_FAILURE: [
                {
                    'strategy': 'wait_and_retry',
                    'priority': 1,
                    'description': 'Wait and retry screenshot capture',
                    'parameters': {'wait_time': 1.0}
                },
                {
                    'strategy': 'restart_service',
                    'priority': 2,
                    'description': 'Restart screenshot service',
                    'parameters': {}
                },
                {
                    'strategy': 'manual_intervention',
                    'priority': 3,
                    'description': 'Request manual intervention',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.TIMEOUT: [
                {
                    'strategy': 'wait_and_retry',
                    'priority': 1,
                    'description': 'Wait longer and retry',
                    'parameters': {'wait_time': 5.0}
                },
                {
                    'strategy': 'adjust_parameters',
                    'priority': 2,
                    'description': 'Increase timeout values',
                    'parameters': {'adjustments': {'timeout': 1.5}}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 3,
                    'description': 'Try faster detection method',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.OCR_ERROR: [
                {
                    'strategy': 'adjust_parameters',
                    'priority': 1,
                    'description': 'Adjust OCR preprocessing parameters',
                    'parameters': {'adjustments': {'ocr_preprocessing': True}}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 2,
                    'description': 'Try template matching instead',
                    'parameters': {}
                },
                {
                    'strategy': 'retry_same_method',
                    'priority': 3,
                    'description': 'Retry OCR with delay',
                    'parameters': {'wait_time': 1.0}
                }
            ],
            
            ErrorCategory.TEMPLATE_ERROR: [
                {
                    'strategy': 'adjust_parameters',
                    'priority': 1,
                    'description': 'Adjust template matching parameters',
                    'parameters': {'adjustments': {'template_threshold': -0.1}}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 2,
                    'description': 'Try OCR instead',
                    'parameters': {}
                },
                {
                    'strategy': 'capture_new_template',
                    'priority': 3,
                    'description': 'Update template images',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.SYSTEM_ERROR: [
                {
                    'strategy': 'wait_and_retry',
                    'priority': 1,
                    'description': 'Wait for system to recover',
                    'parameters': {'wait_time': 3.0}
                },
                {
                    'strategy': 'restart_service',
                    'priority': 2,
                    'description': 'Restart automation service',
                    'parameters': {}
                },
                {
                    'strategy': 'manual_intervention',
                    'priority': 3,
                    'description': 'Request manual intervention',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.CONFIGURATION_ERROR: [
                {
                    'strategy': 'manual_intervention',
                    'priority': 1,
                    'description': 'Fix configuration issue',
                    'parameters': {}
                }
            ],
            
            ErrorCategory.UNKNOWN: [
                {
                    'strategy': 'retry_same_method',
                    'priority': 1,
                    'description': 'Retry with delay',
                    'parameters': {'wait_time': 1.0}
                },
                {
                    'strategy': 'try_next_method',
                    'priority': 2,
                    'description': 'Try alternative method',
                    'parameters': {}
                },
                {
                    'strategy': 'manual_intervention',
                    'priority': 3,
                    'description': 'Request manual review',
                    'parameters': {}
                }
            ]
        }
    
    def _get_parameter_adjustments(self, error_context: ErrorContext) -> Dict[str, Any]:
        """Get parameter adjustments based on error context and history"""
        adjustments = {}
        
        # Check if we have previous adjustments for this element
        element_key = error_context.target_element or 'unknown'
        if element_key in self.parameter_adjustments:
            # Use the most recent successful adjustment
            recent_adjustments = self.parameter_adjustments[element_key][-3:]  # Last 3 attempts
            for adj in recent_adjustments:
                adjustments.update(adj['adjustments'])
        
        return adjustments
    
    def _calculate_strategy_success_rate(self, strategy: RecoveryStrategy, error_context: ErrorContext) -> float:
        """Calculate success rate for a recovery strategy based on history"""
        total_attempts = sum(self.recovery_stats['recovery_strategy_success'].values())
        if total_attempts == 0:
            return 0.5  # Default success rate
        
        successful_attempts = self.recovery_stats['recovery_strategy_success'][strategy.value]
        return successful_attempts / total_attempts if total_attempts > 0 else 0.5
    
    def _update_average_recovery_time(self, recovery_time: float):
        """Update average recovery time"""
        if self.recovery_stats['successful_recoveries'] > 0:
            current_avg = self.recovery_stats['average_recovery_time']
            self.recovery_stats['average_recovery_time'] = (
                (current_avg * (self.recovery_stats['successful_recoveries'] - 1) + recovery_time) / 
                self.recovery_stats['successful_recoveries']
            )
    
    def _get_system_info(self) -> Dict[str, Any]:
        """Get system information for diagnostics"""
        try:
            import platform
            import psutil
            
            return {
                'platform': platform.platform(),
                'python_version': platform.python_version(),
                'cpu_percent': psutil.cpu_percent(),
                'memory_percent': psutil.virtual_memory().percent,
                'disk_usage': psutil.disk_usage('/').percent if os.name != 'nt' else psutil.disk_usage('C:').percent
            }
        except Exception:
            return {'system_info': 'unavailable'}
    
    def _capture_error_screenshot(self, error_context: ErrorContext) -> Optional[str]:
        """Capture screenshot for error analysis"""
        try:
            import pyautogui
            
            timestamp = int(time.time())
            filename = f"error_{error_context.error_category.value}_{timestamp}.png"
            
            debug_dir = self.config.get('debugging.debug_image_path', 'debug_images')
            os.makedirs(debug_dir, exist_ok=True)
            
            screenshot_path = os.path.join(debug_dir, filename)
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            
            return screenshot_path
            
        except Exception as e:
            self.logger.error(f"Failed to capture error screenshot: {e}")
            return None
    
    def _save_error_report(self, error_context: ErrorContext):
        """Save detailed error report to file"""
        try:
            timestamp = int(time.time())
            report_filename = f"error_report_{timestamp}.json"
            
            reports_dir = "EHC_Logs/error_reports"
            os.makedirs(reports_dir, exist_ok=True)
            
            report_path = os.path.join(reports_dir, report_filename)
            
            report_data = {
                'timestamp': error_context.timestamp,
                'error_category': error_context.error_category.value,
                'error_message': error_context.error_message,
                'method_used': error_context.method_used,
                'action_type': error_context.action_type,
                'target_element': error_context.target_element,
                'confidence_score': error_context.confidence_score,
                'screenshot_path': error_context.screenshot_path,
                'stack_trace': error_context.stack_trace,
                'system_info': error_context.system_info,
                'retry_count': error_context.retry_count,
                'previous_attempts': error_context.previous_attempts
            }
            
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump(report_data, f, indent=2, ensure_ascii=False)
            
            self.logger.debug(f"Error report saved: {report_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to save error report: {e}")
    
    def _save_manual_intervention_report(self, error_context: ErrorContext):
        """Save report for manual intervention"""
        try:
            timestamp = int(time.time())
            report_filename = f"manual_intervention_{timestamp}.json"
            
            reports_dir = "EHC_Logs/manual_intervention"
            os.makedirs(reports_dir, exist_ok=True)
            
            report_path = os.path.join(reports_dir, report_filename)
            
            report_data = {
                'timestamp': error_context.timestamp,
                'error_category': error_context.error_category.value,
                'error_message': error_context.error_message,
                'target_element': error_context.target_element,
                'screenshot_path': error_context.screenshot_path,
                'suggested_actions': [
                    "Review screenshot for UI changes",
                    "Update template images if needed",
                    "Adjust detection parameters",
                    "Check system configuration"
                ],
                'system_info': error_context.system_info
            }
            
            with open(report_path, 'w', encoding='utf-8') as f:
                json.dump(report_data, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Manual intervention report saved: {report_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to save manual intervention report: {e}")
    
    def get_error_statistics(self) -> Dict[str, Any]:
        """Get comprehensive error and recovery statistics"""
        stats = self.recovery_stats.copy()
        
        # Calculate success rate
        if stats['total_errors'] > 0:
            stats['recovery_success_rate'] = stats['successful_recoveries'] / stats['total_errors']
        else:
            stats['recovery_success_rate'] = 0.0
        
        # Add recent error trends
        recent_errors = [e for e in self.error_history if time.time() - e.timestamp < 3600]  # Last hour
        stats['recent_error_count'] = len(recent_errors)
        
        # Most common error categories
        stats['most_common_errors'] = sorted(
            stats['error_category_counts'].items(),
            key=lambda x: x[1],
            reverse=True
        )[:5]
        
        return stats
    
    def get_parameter_adjustment_history(self, element_name: str = None) -> Dict[str, Any]:
        """Get parameter adjustment history"""
        if element_name:
            return self.parameter_adjustments.get(element_name, [])
        else:
            return self.parameter_adjustments.copy()
    
    def reset_statistics(self):
        """Reset error and recovery statistics"""
        self.recovery_stats = {
            'total_errors': 0,
            'successful_recoveries': 0,
            'recovery_strategy_success': {strategy.value: 0 for strategy in RecoveryStrategy},
            'error_category_counts': {category.value: 0 for category in ErrorCategory},
            'average_recovery_time': 0.0
        }
        
        self.error_history.clear()
        self.parameter_adjustments.clear()
        
        self.logger.info("Error handler statistics reset")
    
    def export_error_analysis(self, filepath: str = None) -> bool:
        """Export comprehensive error analysis report"""
        try:
            if not filepath:
                timestamp = int(time.time())
                filepath = f"EHC_Logs/error_analysis_{timestamp}.json"
            
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            
            analysis_data = {
                'timestamp': time.time(),
                'statistics': self.get_error_statistics(),
                'parameter_adjustments': self.parameter_adjustment_history,
                'recent_errors': [
                    {
                        'timestamp': e.timestamp,
                        'category': e.error_category.value,
                        'message': e.error_message,
                        'method': e.method_used,
                        'target': e.target_element,
                        'confidence': e.confidence_score
                    }
                    for e in self.error_history[-50:]  # Last 50 errors
                ],
                'recommendations': self._generate_recommendations()
            }
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(analysis_data, f, indent=2, ensure_ascii=False)
            
            self.logger.info(f"Error analysis exported to: {filepath}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to export error analysis: {e}")
            return False
    
    def _generate_recommendations(self) -> List[str]:
        """Generate recommendations based on error patterns"""
        recommendations = []
        
        stats = self.get_error_statistics()
        
        # Check for high error rates
        if stats['recovery_success_rate'] < 0.7:
            recommendations.append("Consider updating template images or adjusting confidence thresholds")
        
        # Check for common error patterns
        most_common = stats['most_common_errors']
        if most_common and most_common[0][1] > 10:
            error_type = most_common[0][0]
            if error_type == 'element_not_found':
                recommendations.append("High element detection failures - review UI element definitions")
            elif error_type == 'low_confidence':
                recommendations.append("Frequent low confidence detections - consider lowering thresholds")
            elif error_type == 'template_error':
                recommendations.append("Template matching issues - update template images")
        
        # Check for recent error spikes
        if stats['recent_error_count'] > 20:
            recommendations.append("High recent error rate - check for system or UI changes")
        
        return recommendations

if __name__ == "__main__":
    # Test error handler
    error_handler = CVErrorHandler()
    print("CV Error Handler initialized successfully")
    
    # Test error handling
    try:
        raise Exception("Test error for demonstration")
    except Exception as e:
        context = {
            'method_used': 'ocr',
            'action_type': 'click',
            'target_element': 'test_button',
            'confidence_score': 0.6
        }
        
        result = error_handler.handle_error(e, context)
        print(f"Recovery result: {result}")