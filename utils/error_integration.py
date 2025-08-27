#!/usr/bin/env python3
"""
Error Handling Integration Module
Integrates comprehensive error handling into existing MoonFlower components
"""

import logging
import functools
from typing import Dict, Any, Optional, Callable, List
from pathlib import Path

from .error_handler import (
    ErrorHandler, ErrorCategory, ErrorSeverity, SystemError, 
    handle_error, handle_errors, get_global_error_handler
)
from .recovery_manager import RecoveryManager


class MoonFlowerErrorIntegration:
    """Integration layer for MoonFlower error handling"""
    
    def __init__(self, config_manager=None):
        """Initialize error integration"""
        self.config_manager = config_manager
        self.error_handler = get_global_error_handler(config_manager)
        self.recovery_manager = RecoveryManager(config_manager)
        self.logger = logging.getLogger(__name__)
        
        # Component-specific error mappings
        self.component_mappings = {
            'csv_downloader': {
                'category': ErrorCategory.NETWORK,
                'common_operations': ['selenium_timeout', 'connection_refused', 'ssl_error']
            },
            'excel_generator': {
                'category': ErrorCategory.DATA_PROCESSING,
                'common_operations': ['csv_parse_error', 'excel_generation_error', 'memory_error']
            },
            'vbs_automator': {
                'category': ErrorCategory.APPLICATION,
                'common_operations': ['process_crash', 'window_not_found', 'automation_failure']
            },
            'email_delivery': {
                'category': ErrorCategory.EMAIL,
                'common_operations': ['smtp_connection_error', 'authentication_error', 'attachment_error']
            },
            'orchestrator': {
                'category': ErrorCategory.SYSTEM,
                'common_operations': ['service_failure', 'scheduling_error', 'coordination_failure']
            },
            'windows_service': {
                'category': ErrorCategory.SYSTEM,
                'common_operations': ['service_start_failure', 'service_stop_failure', 'permission_error']
            }
        }
        
        self.logger.info("🔗 MoonFlower Error Integration initialized")
    
    def handle_component_error(self, 
                              component: str,
                              message: str,
                              operation: str = None,
                              exception: Exception = None,
                              context: Dict[str, Any] = None,
                              severity: ErrorSeverity = None,
                              auto_recover: bool = True) -> SystemError:
        """
        Handle error for specific MoonFlower component
        
        Args:
            component: Component name (e.g., 'csv_downloader', 'vbs_automator')
            message: Error message
            operation: Specific operation that failed
            exception: Original exception
            context: Additional context
            severity: Error severity (auto-detected if not provided)
            auto_recover: Whether to attempt automatic recovery
            
        Returns:
            SystemError object
        """
        try:
            # Get component mapping
            component_info = self.component_mappings.get(component, {})
            category = component_info.get('category', ErrorCategory.UNKNOWN)
            
            # Handle the error
            error = self.error_handler.handle_error(
                category=category,
                message=message,
                component=component,
                operation=operation,
                exception=exception,
                context=context or {},
                severity=severity,
                auto_recover=auto_recover
            )
            
            # Attempt component-specific recovery if auto_recover is enabled
            if auto_recover and not error.resolved:
                recovery_success = self._attempt_component_recovery(error)
                if recovery_success:
                    error.mark_resolved()
                    self.logger.info(f"✅ Component recovery successful for {component}")
            
            return error
            
        except Exception as e:
            self.logger.error(f"Error integration failed: {e}")
            # Return a basic error if integration fails
            return SystemError(
                category=ErrorCategory.SYSTEM,
                severity=ErrorSeverity.CRITICAL,
                message=f"Error integration failure: {str(e)}",
                component="error_integration",
                exception=e
            )
    
    def _attempt_component_recovery(self, error: SystemError) -> bool:
        """Attempt component-specific recovery"""
        try:
            component = error.component
            category = error.category
            
            # Route to appropriate recovery method
            if category == ErrorCategory.NETWORK:
                return self.recovery_manager.recover_network_issues(error)
            elif category == ErrorCategory.APPLICATION:
                return self.recovery_manager.recover_application_crashes(error)
            elif category == ErrorCategory.FILE_SYSTEM:
                return self.recovery_manager.recover_file_system_issues(error)
            elif category == ErrorCategory.DATA_PROCESSING:
                return self.recovery_manager.recover_data_processing_issues(error)
            elif category == ErrorCategory.EMAIL:
                return self.recovery_manager.recover_email_issues(error)
            elif category == ErrorCategory.SYSTEM:
                return self.recovery_manager.recover_system_issues(error)
            else:
                return False
                
        except Exception as e:
            self.logger.error(f"Component recovery failed: {e}")
            return False
    
    def create_component_decorator(self, component: str, operation: str = None):
        """Create error handling decorator for specific component"""
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    # Handle the error
                    error = self.handle_component_error(
                        component=component,
                        message=f"Function {func.__name__} failed: {str(e)}",
                        operation=operation or func.__name__,
                        exception=e,
                        context={
                            "function": func.__name__,
                            "module": func.__module__,
                            "args_count": len(args),
                            "kwargs_keys": list(kwargs.keys())
                        }
                    )
                    
                    # Re-raise the exception after handling
                    raise
            return wrapper
        return decorator
    
    def get_component_health_status(self) -> Dict[str, Any]:
        """Get health status for all components"""
        try:
            summary = self.error_handler.get_error_summary()
            recovery_stats = self.recovery_manager.get_recovery_stats()
            
            # Analyze component health
            component_health = {}
            
            for component in self.component_mappings.keys():
                component_errors = [
                    e for e in self.error_handler.errors 
                    if e.component == component
                ]
                
                recent_errors = [
                    e for e in component_errors 
                    if (e.timestamp - e.timestamp).total_seconds() < 3600  # Last hour
                ]
                
                unresolved_errors = [e for e in component_errors if not e.resolved]
                
                # Determine health status
                if len(unresolved_errors) > 5:
                    health_status = "critical"
                elif len(recent_errors) > 10:
                    health_status = "degraded"
                elif len(recent_errors) > 0:
                    health_status = "warning"
                else:
                    health_status = "healthy"
                
                component_health[component] = {
                    "status": health_status,
                    "total_errors": len(component_errors),
                    "recent_errors": len(recent_errors),
                    "unresolved_errors": len(unresolved_errors),
                    "last_error": component_errors[-1].to_dict() if component_errors else None
                }
            
            return {
                "overall_summary": summary,
                "recovery_stats": recovery_stats,
                "component_health": component_health,
                "system_recommendations": self._generate_system_recommendations(component_health)
            }
            
        except Exception as e:
            return {"error": f"Health status generation failed: {e}"}
    
    def _generate_system_recommendations(self, component_health: Dict[str, Any]) -> List[str]:
        """Generate system recommendations based on component health"""
        recommendations = []
        
        try:
            # Check for critical components
            critical_components = [
                comp for comp, health in component_health.items()
                if health["status"] == "critical"
            ]
            
            if critical_components:
                recommendations.append(
                    f"🚨 Critical components need immediate attention: {', '.join(critical_components)}"
                )
            
            # Check for degraded components
            degraded_components = [
                comp for comp, health in component_health.items()
                if health["status"] == "degraded"
            ]
            
            if degraded_components:
                recommendations.append(
                    f"⚠️ Degraded components may need maintenance: {', '.join(degraded_components)}"
                )
            
            # Check for high error rates
            high_error_components = [
                comp for comp, health in component_health.items()
                if health["recent_errors"] > 5
            ]
            
            if high_error_components:
                recommendations.append(
                    f"📊 High error rate detected in: {', '.join(high_error_components)}"
                )
            
            # System-wide recommendations
            total_unresolved = sum(
                health["unresolved_errors"] for health in component_health.values()
            )
            
            if total_unresolved > 20:
                recommendations.append("🔄 Consider system restart to clear accumulated errors")
            
            if not recommendations:
                recommendations.append("✅ All components are operating normally")
            
        except Exception as e:
            recommendations.append(f"❌ Failed to generate recommendations: {e}")
        
        return recommendations
    
    def create_error_context(self, **kwargs) -> Dict[str, Any]:
        """Create standardized error context"""
        context = {
            "timestamp": kwargs.get("timestamp", ""),
            "user_session": kwargs.get("user_session", "system"),
            "system_info": {
                "python_version": kwargs.get("python_version", ""),
                "platform": kwargs.get("platform", ""),
                "memory_usage": kwargs.get("memory_usage", 0),
                "cpu_usage": kwargs.get("cpu_usage", 0)
            }
        }
        
        # Add any additional context
        for key, value in kwargs.items():
            if key not in ["timestamp", "user_session", "python_version", "platform", "memory_usage", "cpu_usage"]:
                context[key] = value
        
        return context


# Component-specific error handling functions
def handle_csv_downloader_error(message: str, 
                               operation: str = None,
                               exception: Exception = None,
                               context: Dict[str, Any] = None,
                               integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle CSV downloader specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="csv_downloader",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


def handle_excel_generator_error(message: str,
                                operation: str = None,
                                exception: Exception = None,
                                context: Dict[str, Any] = None,
                                integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle Excel generator specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="excel_generator",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


def handle_vbs_automator_error(message: str,
                              operation: str = None,
                              exception: Exception = None,
                              context: Dict[str, Any] = None,
                              integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle VBS automator specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="vbs_automator",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


def handle_email_delivery_error(message: str,
                               operation: str = None,
                               exception: Exception = None,
                               context: Dict[str, Any] = None,
                               integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle email delivery specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="email_delivery",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


def handle_orchestrator_error(message: str,
                             operation: str = None,
                             exception: Exception = None,
                             context: Dict[str, Any] = None,
                             integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle orchestrator specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="orchestrator",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


def handle_service_error(message: str,
                        operation: str = None,
                        exception: Exception = None,
                        context: Dict[str, Any] = None,
                        integration: MoonFlowerErrorIntegration = None) -> SystemError:
    """Handle Windows service specific errors"""
    if integration is None:
        integration = MoonFlowerErrorIntegration()
    
    return integration.handle_component_error(
        component="windows_service",
        message=message,
        operation=operation,
        exception=exception,
        context=context
    )


# Decorator factories for each component
def csv_downloader_errors(operation: str = None):
    """Decorator for CSV downloader error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("csv_downloader", operation)


def excel_generator_errors(operation: str = None):
    """Decorator for Excel generator error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("excel_generator", operation)


def vbs_automator_errors(operation: str = None):
    """Decorator for VBS automator error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("vbs_automator", operation)


def email_delivery_errors(operation: str = None):
    """Decorator for email delivery error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("email_delivery", operation)


def orchestrator_errors(operation: str = None):
    """Decorator for orchestrator error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("orchestrator", operation)


def service_errors(operation: str = None):
    """Decorator for Windows service error handling"""
    integration = MoonFlowerErrorIntegration()
    return integration.create_component_decorator("windows_service", operation)


# Global integration instance
_global_integration = None

def get_global_integration(config_manager=None) -> MoonFlowerErrorIntegration:
    """Get or create global integration instance"""
    global _global_integration
    if _global_integration is None:
        _global_integration = MoonFlowerErrorIntegration(config_manager)
    return _global_integration


if __name__ == "__main__":
    # Test the integration
    print("🧪 Testing MoonFlower Error Integration")
    print("=" * 60)
    
    integration = MoonFlowerErrorIntegration()
    
    # Test component error handling
    error = integration.handle_component_error(
        component="csv_downloader",
        message="Test network timeout",
        operation="selenium_timeout",
        context={"url": "https://example.com", "timeout": 30}
    )
    
    print(f"Handled error: {error.id} - {error.message}")
    
    # Test decorator
    @csv_downloader_errors("test_operation")
    def test_function():
        raise ValueError("Test exception")
    
    try:
        test_function()
    except ValueError:
        print("Exception properly handled and re-raised")
    
    # Test health status
    health_status = integration.get_component_health_status()
    print(f"\nComponent Health Status:")
    for component, health in health_status["component_health"].items():
        print(f"  {component}: {health['status']} ({health['total_errors']} total errors)")
    
    print(f"\nRecommendations:")
    for recommendation in health_status["system_recommendations"]:
        print(f"  {recommendation}")
    
    print("\nIntegration test completed")