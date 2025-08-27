#!/usr/bin/env python3
"""
VBS Error Handler Module
Comprehensive error handling and recovery for VBS automation failures
"""

import time
import logging
import win32gui
import win32con
import win32api
import win32process
import psutil
import os
from typing import Dict, Optional, Any, List, Callable
import traceback
from datetime import datetime
from enum import Enum

class VBSErrorType(Enum):
    """VBS Error Types"""
    WINDOW_NOT_FOUND = "window_not_found"
    WINDOW_FOCUS_FAILED = "window_focus_failed"
    COORDINATE_CLICK_FAILED = "coordinate_click_failed"
    KEY_PRESS_FAILED = "key_press_failed"
    FILE_NOT_FOUND = "file_not_found"
    FILE_DIALOG_FAILED = "file_dialog_failed"
    TIMEOUT_EXCEEDED = "timeout_exceeded"
    PROCESS_CRASHED = "process_crashed"
    NETWORK_ERROR = "network_error"
    PERMISSION_DENIED = "permission_denied"
    UNKNOWN_ERROR = "unknown_error"

class VBSErrorSeverity(Enum):
    """VBS Error Severity Levels"""
    LOW = "low"           # Recoverable, retry possible
    MEDIUM = "medium"     # Requires intervention, may recover
    HIGH = "high"         # Critical, requires restart
    CRITICAL = "critical" # Fatal, requires manual intervention

class VBSError:
    """VBS Error Information"""
    
    def __init__(self, error_type: VBSErrorType, severity: VBSErrorSeverity, 
                 message: str, phase: str = None, step: str = None, 
                 exception: Exception = None, context: Dict[str, Any] = None):
        self.error_type = error_type
        self.severity = severity
        self.message = message
        self.phase = phase
        self.step = step
        self.exception = exception
        self.context = context or {}
        self.timestamp = datetime.now()
        self.retry_count = 0
        self.resolved = False
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert error to dictionary"""
        return {
            "error_type": self.error_type.value,
            "severity": self.severity.value,
            "message": self.message,
            "phase": self.phase,
            "step": self.step,
            "timestamp": self.timestamp.isoformat(),
            "retry_count": self.retry_count,
            "resolved": self.resolved,
            "context": self.context,
            "exception": str(self.exception) if self.exception else None
        }

class VBSErrorHandler:
    """Comprehensive VBS Error Handler"""
    
    def __init__(self, max_retries: int = 3, retry_delay: float = 2.0):
        """Initialize VBS Error Handler"""
        self.logger = self._setup_logging()
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        
        # Error tracking
        self.errors: List[VBSError] = []
        self.error_counts = {}
        self.recovery_strategies = {}
        
        # Recovery callbacks
        self.recovery_callbacks = {
            VBSErrorType.WINDOW_NOT_FOUND: self._recover_window_not_found,
            VBSErrorType.WINDOW_FOCUS_FAILED: self._recover_window_focus_failed,
            VBSErrorType.COORDINATE_CLICK_FAILED: self._recover_coordinate_click_failed,
            VBSErrorType.KEY_PRESS_FAILED: self._recover_key_press_failed,
            VBSErrorType.FILE_NOT_FOUND: self._recover_file_not_found,
            VBSErrorType.FILE_DIALOG_FAILED: self._recover_file_dialog_failed,
            VBSErrorType.TIMEOUT_EXCEEDED: self._recover_timeout_exceeded,
            VBSErrorType.PROCESS_CRASHED: self._recover_process_crashed,
        }
        
        self.logger.info("🛡️ VBS Error Handler initialized")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup enhanced logging"""
        logger = logging.getLogger("VBSErrorHandler")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            try:
                log_file = "EHC_Logs/vbs_error_handler.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def handle_error(self, error_type: VBSErrorType, message: str, 
                    phase: str = None, step: str = None, 
                    exception: Exception = None, context: Dict[str, Any] = None,
                    severity: VBSErrorSeverity = None) -> VBSError:
        """Handle VBS error with automatic severity detection and recovery"""
        try:
            # Auto-detect severity if not provided
            if severity is None:
                severity = self._detect_error_severity(error_type, exception, context)
            
            # Create error object
            error = VBSError(
                error_type=error_type,
                severity=severity,
                message=message,
                phase=phase,
                step=step,
                exception=exception,
                context=context
            )
            
            # Log error
            self._log_error(error)
            
            # Track error
            self.errors.append(error)
            self._update_error_counts(error_type)
            
            # Attempt recovery if possible
            if severity in [VBSErrorSeverity.LOW, VBSErrorSeverity.MEDIUM]:
                recovery_success = self._attempt_recovery(error)
                if recovery_success:
                    error.resolved = True
                    self.logger.info(f"✅ Error recovered: {error_type.value}")
            
            return error
            
        except Exception as e:
            self.logger.error(f"❌ Error handler failed: {e}")
            # Return a basic error if error handling itself fails
            return VBSError(
                error_type=VBSErrorType.UNKNOWN_ERROR,
                severity=VBSErrorSeverity.CRITICAL,
                message=f"Error handler failure: {str(e)}",
                exception=e
            )
    
    def _detect_error_severity(self, error_type: VBSErrorType, 
                              exception: Exception = None, 
                              context: Dict[str, Any] = None) -> VBSErrorSeverity:
        """Automatically detect error severity"""
        try:
            # Critical errors that require manual intervention
            critical_errors = [
                VBSErrorType.PROCESS_CRASHED,
                VBSErrorType.PERMISSION_DENIED
            ]
            
            # High severity errors that require restart
            high_errors = [
                VBSErrorType.WINDOW_NOT_FOUND,
                VBSErrorType.TIMEOUT_EXCEEDED
            ]
            
            # Medium severity errors that may recover
            medium_errors = [
                VBSErrorType.WINDOW_FOCUS_FAILED,
                VBSErrorType.FILE_DIALOG_FAILED,
                VBSErrorType.NETWORK_ERROR
            ]
            
            # Low severity errors that are easily recoverable
            low_errors = [
                VBSErrorType.COORDINATE_CLICK_FAILED,
                VBSErrorType.KEY_PRESS_FAILED,
                VBSErrorType.FILE_NOT_FOUND
            ]
            
            if error_type in critical_errors:
                return VBSErrorSeverity.CRITICAL
            elif error_type in high_errors:
                return VBSErrorSeverity.HIGH
            elif error_type in medium_errors:
                return VBSErrorSeverity.MEDIUM
            elif error_type in low_errors:
                return VBSErrorSeverity.LOW
            else:
                return VBSErrorSeverity.MEDIUM  # Default
                
        except Exception:
            return VBSErrorSeverity.MEDIUM  # Safe default
    
    def _log_error(self, error: VBSError):
        """Log error with appropriate level"""
        try:
            error_msg = f"🚨 VBS Error [{error.severity.value.upper()}]: {error.message}"
            if error.phase:
                error_msg += f" (Phase: {error.phase})"
            if error.step:
                error_msg += f" (Step: {error.step})"
            
            # Log based on severity
            if error.severity == VBSErrorSeverity.CRITICAL:
                self.logger.critical(error_msg)
            elif error.severity == VBSErrorSeverity.HIGH:
                self.logger.error(error_msg)
            elif error.severity == VBSErrorSeverity.MEDIUM:
                self.logger.warning(error_msg)
            else:
                self.logger.info(error_msg)
            
            # Log exception details if available
            if error.exception:
                self.logger.debug(f"Exception details: {traceback.format_exc()}")
                
        except Exception as e:
            print(f"Logging failed: {e}")  # Fallback to print
    
    def _update_error_counts(self, error_type: VBSErrorType):
        """Update error occurrence counts"""
        try:
            if error_type not in self.error_counts:
                self.error_counts[error_type] = 0
            self.error_counts[error_type] += 1
            
            # Log frequent errors
            if self.error_counts[error_type] > 5:
                self.logger.warning(f"⚠️ Frequent error detected: {error_type.value} ({self.error_counts[error_type]} times)")
                
        except Exception as e:
            self.logger.error(f"❌ Error count update failed: {e}")
    
    def _attempt_recovery(self, error: VBSError) -> bool:
        """Attempt to recover from error"""
        try:
            if error.retry_count >= self.max_retries:
                self.logger.warning(f"⚠️ Max retries exceeded for {error.error_type.value}")
                return False
            
            error.retry_count += 1
            self.logger.info(f"🔄 Attempting recovery for {error.error_type.value} (attempt {error.retry_count})")
            
            # Get recovery callback
            recovery_callback = self.recovery_callbacks.get(error.error_type)
            if not recovery_callback:
                self.logger.warning(f"⚠️ No recovery strategy for {error.error_type.value}")
                return False
            
            # Wait before retry
            time.sleep(self.retry_delay * error.retry_count)
            
            # Attempt recovery
            recovery_success = recovery_callback(error)
            
            if recovery_success:
                self.logger.info(f"✅ Recovery successful for {error.error_type.value}")
            else:
                self.logger.warning(f"❌ Recovery failed for {error.error_type.value}")
            
            return recovery_success
            
        except Exception as e:
            self.logger.error(f"❌ Recovery attempt failed: {e}")
            return False
    
    def _recover_window_not_found(self, error: VBSError) -> bool:
        """Recover from window not found error"""
        try:
            self.logger.info("🔍 Attempting to find VBS window...")
            
            # Search for VBS windows
            vbs_windows = []
            
            def enum_windows_callback(hwnd, windows):
                try:
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        if title:
                            title_lower = title.lower()
                            vbs_indicators = ['absons', 'arabian', 'moonflower', 'erp', 'wifi', 'user']
                            exclude_indicators = ['login', 'security', 'warning', 'browser']
                            
                            has_vbs = any(indicator in title_lower for indicator in vbs_indicators)
                            has_exclude = any(indicator in title_lower for indicator in exclude_indicators)
                            
                            if has_vbs and not has_exclude:
                                windows.append((hwnd, title))
                except:
                    pass
                return True
            
            win32gui.EnumWindows(enum_windows_callback, vbs_windows)
            
            if vbs_windows:
                self.logger.info(f"✅ Found VBS window: {vbs_windows[0][1]}")
                # Update context with found window
                error.context['recovered_window_handle'] = vbs_windows[0][0]
                return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Window recovery failed: {e}")
            return False
    
    def _recover_window_focus_failed(self, error: VBSError) -> bool:
        """Recover from window focus failure"""
        try:
            window_handle = error.context.get('window_handle')
            if not window_handle:
                return False
            
            self.logger.info("🎯 Attempting to restore window focus...")
            
            # Try multiple focus methods
            try:
                # Method 1: Standard focus
                win32gui.SetForegroundWindow(window_handle)
                win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
                time.sleep(0.5)
                
                # Verify focus
                focused_window = win32gui.GetForegroundWindow()
                if focused_window == window_handle:
                    return True
                
                # Method 2: Force focus with Alt+Tab simulation
                win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)  # Alt down
                win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)   # Tab down
                time.sleep(0.1)
                win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)   # Tab up
                win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)  # Alt up
                
                time.sleep(0.5)
                return True
                
            except Exception as e:
                self.logger.warning(f"⚠️ Focus recovery method failed: {e}")
                return False
            
        except Exception as e:
            self.logger.error(f"❌ Window focus recovery failed: {e}")
            return False
    
    def _recover_coordinate_click_failed(self, error: VBSError) -> bool:
        """Recover from coordinate click failure"""
        try:
            self.logger.info("🖱️ Attempting coordinate click recovery...")
            
            coordinate = error.context.get('coordinate')
            window_handle = error.context.get('window_handle')
            
            if not coordinate or not window_handle:
                return False
            
            # Try alternative click methods
            x, y = coordinate
            
            # Get window position
            window_rect = win32gui.GetWindowRect(window_handle)
            screen_x = window_rect[0] + x
            screen_y = window_rect[1] + y
            
            # Method 1: Double click
            win32api.SetCursorPos((screen_x, screen_y))
            time.sleep(0.2)
            win32api.mouse_event(win32api.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.1)
            win32api.mouse_event(win32api.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            time.sleep(0.1)
            win32api.mouse_event(win32api.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            time.sleep(0.1)
            win32api.mouse_event(win32api.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Coordinate click recovery failed: {e}")
            return False
    
    def _recover_key_press_failed(self, error: VBSError) -> bool:
        """Recover from key press failure"""
        try:
            self.logger.info("⌨️ Attempting key press recovery...")
            
            key_name = error.context.get('key_name')
            window_handle = error.context.get('window_handle')
            
            if not key_name:
                return False
            
            # Key mapping
            key_codes = {
                "ENTER": win32con.VK_RETURN,
                "TAB": win32con.VK_TAB,
                "LEFT": win32con.VK_LEFT,
                "RIGHT": win32con.VK_RIGHT,
                "ESCAPE": win32con.VK_ESCAPE
            }
            
            if key_name not in key_codes:
                return False
            
            vk_code = key_codes[key_name]
            
            # Try multiple key press methods
            try:
                # Method 1: Global key event
                win32api.keybd_event(vk_code, 0, 0, 0)
                time.sleep(0.1)
                win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)
                return True
                
            except Exception:
                # Method 2: Window message (if window handle available)
                if window_handle:
                    win32api.SendMessage(window_handle, win32con.WM_KEYDOWN, vk_code, 0)
                    time.sleep(0.1)
                    win32api.SendMessage(window_handle, win32con.WM_KEYUP, vk_code, 0)
                    return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Key press recovery failed: {e}")
            return False
    
    def _recover_file_not_found(self, error: VBSError) -> bool:
        """Recover from file not found error"""
        try:
            self.logger.info("📁 Attempting file recovery...")
            
            file_path = error.context.get('file_path')
            if not file_path:
                return False
            
            # Try to find alternative files
            from pathlib import Path
            
            file_path_obj = Path(file_path)
            parent_dir = file_path_obj.parent
            file_pattern = file_path_obj.stem + "*" + file_path_obj.suffix
            
            if parent_dir.exists():
                matching_files = list(parent_dir.glob(file_pattern))
                if matching_files:
                    # Use the most recent file
                    latest_file = max(matching_files, key=os.path.getctime)
                    error.context['recovered_file_path'] = str(latest_file)
                    self.logger.info(f"✅ Found alternative file: {latest_file}")
                    return True
            
            return False
            
        except Exception as e:
            self.logger.error(f"❌ File recovery failed: {e}")
            return False
    
    def _recover_file_dialog_failed(self, error: VBSError) -> bool:
        """Recover from file dialog failure"""
        try:
            self.logger.info("📂 Attempting file dialog recovery...")
            
            # Try to close any open dialogs
            win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
            
            time.sleep(1.0)
            
            # Try Alt+F4 to close dialog
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
            win32api.keybd_event(win32con.VK_F4, 0, 0, 0)
            time.sleep(0.1)
            win32api.keybd_event(win32con.VK_F4, 0, win32con.KEYEVENTF_KEYUP, 0)
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            
            time.sleep(1.0)
            return True
            
        except Exception as e:
            self.logger.error(f"❌ File dialog recovery failed: {e}")
            return False
    
    def _recover_timeout_exceeded(self, error: VBSError) -> bool:
        """Recover from timeout exceeded error"""
        try:
            self.logger.info("⏰ Attempting timeout recovery...")
            
            # For timeout errors, we typically need to restart the operation
            # This is more of a notification than a recovery
            timeout_type = error.context.get('timeout_type', 'unknown')
            
            if timeout_type == 'update_process':
                # For update process timeouts, we might want to check if it actually completed
                self.logger.info("🔍 Checking if update process completed despite timeout...")
                # This would require specific logic to check update completion
                return False  # Cannot automatically recover from update timeout
            
            return False  # Most timeouts require manual intervention
            
        except Exception as e:
            self.logger.error(f"❌ Timeout recovery failed: {e}")
            return False
    
    def _recover_process_crashed(self, error: VBSError) -> bool:
        """Recover from process crash"""
        try:
            self.logger.info("💥 Attempting process crash recovery...")
            
            # Process crashes typically require restart
            # This is more of a cleanup than recovery
            process_id = error.context.get('process_id')
            
            if process_id:
                try:
                    process = psutil.Process(process_id)
                    if process.is_running():
                        process.terminate()
                        process.wait(timeout=10)
                except:
                    pass
            
            # Cannot automatically recover from process crash
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Process crash recovery failed: {e}")
            return False
    
    def get_error_summary(self) -> Dict[str, Any]:
        """Get comprehensive error summary"""
        try:
            total_errors = len(self.errors)
            resolved_errors = sum(1 for error in self.errors if error.resolved)
            
            severity_counts = {}
            for severity in VBSErrorSeverity:
                severity_counts[severity.value] = sum(
                    1 for error in self.errors if error.severity == severity
                )
            
            type_counts = {}
            for error_type in VBSErrorType:
                type_counts[error_type.value] = self.error_counts.get(error_type, 0)
            
            recent_errors = [
                error.to_dict() for error in self.errors[-10:]  # Last 10 errors
            ]
            
            return {
                "total_errors": total_errors,
                "resolved_errors": resolved_errors,
                "unresolved_errors": total_errors - resolved_errors,
                "resolution_rate": (resolved_errors / total_errors * 100) if total_errors > 0 else 0,
                "severity_breakdown": severity_counts,
                "error_type_breakdown": type_counts,
                "recent_errors": recent_errors,
                "max_retries": self.max_retries,
                "retry_delay": self.retry_delay
            }
            
        except Exception as e:
            self.logger.error(f"❌ Error summary generation failed: {e}")
            return {"error": str(e)}
    
    def clear_resolved_errors(self):
        """Clear resolved errors from tracking"""
        try:
            initial_count = len(self.errors)
            self.errors = [error for error in self.errors if not error.resolved]
            cleared_count = initial_count - len(self.errors)
            
            if cleared_count > 0:
                self.logger.info(f"🧹 Cleared {cleared_count} resolved errors")
                
        except Exception as e:
            self.logger.error(f"❌ Error clearing failed: {e}")

# Convenience functions for easy error handling
def handle_vbs_error(error_type: VBSErrorType, message: str, 
                    phase: str = None, step: str = None,
                    exception: Exception = None, context: Dict[str, Any] = None,
                    error_handler: VBSErrorHandler = None) -> VBSError:
    """Handle VBS error with global error handler"""
    if error_handler is None:
        error_handler = VBSErrorHandler()
    
    return error_handler.handle_error(
        error_type=error_type,
        message=message,
        phase=phase,
        step=step,
        exception=exception,
        context=context
    )

def create_error_context(window_handle: int = None, coordinate: tuple = None,
                        file_path: str = None, key_name: str = None,
                        timeout_type: str = None, process_id: int = None,
                        **kwargs) -> Dict[str, Any]:
    """Create error context dictionary"""
    context = {}
    
    if window_handle is not None:
        context['window_handle'] = window_handle
    if coordinate is not None:
        context['coordinate'] = coordinate
    if file_path is not None:
        context['file_path'] = file_path
    if key_name is not None:
        context['key_name'] = key_name
    if timeout_type is not None:
        context['timeout_type'] = timeout_type
    if process_id is not None:
        context['process_id'] = process_id
    
    # Add any additional context
    context.update(kwargs)
    
    return context

if __name__ == "__main__":
    # Test the error handler
    print("🧪 Testing VBS Error Handler")
    print("=" * 50)
    
    error_handler = VBSErrorHandler()
    
    # Test different error types
    test_errors = [
        (VBSErrorType.WINDOW_NOT_FOUND, "Test window not found"),
        (VBSErrorType.COORDINATE_CLICK_FAILED, "Test click failed"),
        (VBSErrorType.FILE_NOT_FOUND, "Test file missing"),
    ]
    
    for error_type, message in test_errors:
        error = error_handler.handle_error(
            error_type=error_type,
            message=message,
            phase="test_phase",
            context=create_error_context(window_handle=12345)
        )
        print(f"Handled error: {error.error_type.value} - {error.message}")
    
    # Print summary
    summary = error_handler.get_error_summary()
    print(f"\nError Summary:")
    print(f"  Total errors: {summary['total_errors']}")
    print(f"  Resolved: {summary['resolved_errors']}")
    print(f"  Resolution rate: {summary['resolution_rate']:.1f}%")
    
    print("Test completed")