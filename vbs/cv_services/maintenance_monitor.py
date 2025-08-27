#!/usr/bin/env python3
"""
Maintenance and Monitoring System for VBS Computer Vision Services
Provides automated template validation, health checks, performance monitoring, and backup/recovery
"""

import os
import json
import time
import logging
import shutil
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, asdict
from pathlib import Path
import cv2
import numpy as np
from .config_loader import get_cv_config
from .template_service import TemplateService
from .ocr_service import OCRService
from .smart_engine import SmartAutomationEngine

@dataclass
class HealthCheckResult:
    """Result of a system health check"""
    component: str
    status: str  # healthy, warning, critical
    message: str
    timestamp: float
    metrics: Dict[str, Any]

@dataclass
class PerformanceMetrics:
    """Performance metrics for monitoring"""
    timestamp: float
    component: str
    operation_count: int
    success_rate: float
    average_response_time: float
    error_count: int
    memory_usage_mb: float
    cpu_usage_percent: float

@dataclass
class MaintenanceTask:
    """Represents a maintenance task"""
    task_id: str
    name: str
    description: str
    frequency_hours: int
    last_run: float
    next_run: float
    enabled: bool
    task_function: str

class MaintenanceMonitor:
    """Main maintenance and monitoring system"""
    
    def __init__(self):
        self.config = get_cv_config()
        self.logger = self._setup_logging()
        self.maintenance_config = self.config.get('maintenance', {})
        
        # Initialize services for monitoring
        self.template_service = TemplateService()
        self.ocr_service = OCRService()
        self.smart_engine = SmartAutomationEngine()
        
        # Monitoring state
        self.monitoring_active = False
        self.monitoring_thread = None
        self.health_history = []
        self.performance_history = []
        self.error_reports = []
        
        # Maintenance tasks
        self.maintenance_tasks = self._initialize_maintenance_tasks()
        
        # Backup settings
        self.backup_config = self.maintenance_config.get('backup', {})
        self.backup_directory = self.backup_config.get('directory', 'backups/cv_system')
        self.max_backups = self.backup_config.get('max_backups', 10)
        
        self.logger.info("Maintenance Monitor initialized successfully")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for maintenance monitor"""
        logger = logging.getLogger("MaintenanceMonitor")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler
            try:
                log_file = "EHC_Logs/maintenance_monitor.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def _initialize_maintenance_tasks(self) -> List[MaintenanceTask]:
        """Initialize maintenance tasks"""
        tasks = [
            MaintenanceTask(
                task_id="template_validation",
                name="Template Validation",
                description="Validate all templates and check for outdated ones",
                frequency_hours=24,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="validate_templates"
            ),
            MaintenanceTask(
                task_id="performance_analysis",
                name="Performance Analysis",
                description="Analyze performance metrics and generate recommendations",
                frequency_hours=12,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="analyze_performance"
            ),
            MaintenanceTask(
                task_id="error_analysis",
                name="Error Analysis",
                description="Analyze error patterns and generate reports",
                frequency_hours=6,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="analyze_errors"
            ),
            MaintenanceTask(
                task_id="system_backup",
                name="System Backup",
                description="Backup templates, configuration, and performance data",
                frequency_hours=24,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="create_system_backup"
            ),
            MaintenanceTask(
                task_id="health_check",
                name="System Health Check",
                description="Comprehensive system health assessment",
                frequency_hours=1,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="perform_health_check"
            ),
            MaintenanceTask(
                task_id="cleanup_old_data",
                name="Data Cleanup",
                description="Clean up old logs, debug images, and temporary files",
                frequency_hours=48,
                last_run=0,
                next_run=time.time(),
                enabled=True,
                task_function="cleanup_old_data"
            )
        ]
        
        return tasks
    
    def start_monitoring(self):
        """Start the monitoring system"""
        if self.monitoring_active:
            self.logger.warning("Monitoring already active")
            return
        
        self.monitoring_active = True
        self.monitoring_thread = threading.Thread(target=self._monitoring_loop, daemon=True)
        self.monitoring_thread.start()
        
        self.logger.info("Maintenance monitoring started")
    
    def stop_monitoring(self):
        """Stop the monitoring system"""
        self.monitoring_active = False
        if self.monitoring_thread:
            self.monitoring_thread.join(timeout=5)
        
        self.logger.info("Maintenance monitoring stopped")
    
    def _monitoring_loop(self):
        """Main monitoring loop"""
        while self.monitoring_active:
            try:
                current_time = time.time()
                
                # Check for due maintenance tasks
                for task in self.maintenance_tasks:
                    if task.enabled and current_time >= task.next_run:
                        self._execute_maintenance_task(task)
                
                # Collect performance metrics
                self._collect_performance_metrics()
                
                # Sleep for monitoring interval
                monitoring_interval = self.maintenance_config.get('monitoring_interval_seconds', 60)
                time.sleep(monitoring_interval)
                
            except Exception as e:
                self.logger.error(f"Error in monitoring loop: {e}")
                time.sleep(60)  # Wait before retrying
    
    def _execute_maintenance_task(self, task: MaintenanceTask):
        """Execute a maintenance task"""
        try:
            self.logger.info(f"Executing maintenance task: {task.name}")
            
            # Get task function
            task_func = getattr(self, task.task_function, None)
            if not task_func:
                self.logger.error(f"Task function not found: {task.task_function}")
                return
            
            # Execute task
            start_time = time.time()
            result = task_func()
            execution_time = time.time() - start_time
            
            # Update task schedule
            task.last_run = time.time()
            task.next_run = task.last_run + (task.frequency_hours * 3600)
            
            self.logger.info(f"Maintenance task completed: {task.name} (took {execution_time:.2f}s)")
            
            # Log task result
            if result:
                self._log_maintenance_result(task, result, execution_time)
            
        except Exception as e:
            self.logger.error(f"Error executing maintenance task {task.name}: {e}")
    
    def validate_templates(self) -> Dict[str, Any]:
        """Validate all templates and identify issues"""
        try:
            self.logger.info("Starting template validation")
            
            validation_results = {
                'timestamp': time.time(),
                'total_templates': 0,
                'valid_templates': 0,
                'outdated_templates': [],
                'low_performance_templates': [],
                'missing_metadata': [],
                'recommendations': []
            }
            
            template_list = self.template_service.get_template_list()
            validation_results['total_templates'] = len(template_list)
            
            for template_name in template_list:
                try:
                    # Get template info
                    template_info = self.template_service.get_template_info(template_name)
                    if not template_info:
                        validation_results['missing_metadata'].append(template_name)
                        continue
                    
                    # Check if template is outdated (not used in 30 days)
                    last_used = template_info.get('last_used', 0)
                    if time.time() - last_used > (30 * 24 * 3600):
                        validation_results['outdated_templates'].append({
                            'name': template_name,
                            'last_used': last_used,
                            'days_unused': (time.time() - last_used) / (24 * 3600)
                        })
                    
                    # Check performance
                    runtime_stats = template_info.get('runtime_stats', {})
                    success_rate = runtime_stats.get('current_success_rate', 0)
                    if success_rate < 0.5 and runtime_stats.get('total_usage', 0) > 10:
                        validation_results['low_performance_templates'].append({
                            'name': template_name,
                            'success_rate': success_rate,
                            'usage_count': runtime_stats.get('total_usage', 0)
                        })
                    
                    validation_results['valid_templates'] += 1
                    
                except Exception as e:
                    self.logger.warning(f"Error validating template {template_name}: {e}")
            
            # Generate recommendations
            if validation_results['outdated_templates']:
                validation_results['recommendations'].append(
                    f"Consider removing {len(validation_results['outdated_templates'])} outdated templates"
                )
            
            if validation_results['low_performance_templates']:
                validation_results['recommendations'].append(
                    f"Update {len(validation_results['low_performance_templates'])} low-performance templates"
                )
            
            if validation_results['missing_metadata']:
                validation_results['recommendations'].append(
                    f"Add metadata for {len(validation_results['missing_metadata'])} templates"
                )
            
            self.logger.info(f"Template validation completed: {validation_results['valid_templates']}/{validation_results['total_templates']} valid")
            
            return validation_results
            
        except Exception as e:
            self.logger.error(f"Template validation failed: {e}")
            return {'error': str(e)}
    
    def analyze_performance(self) -> Dict[str, Any]:
        """Analyze performance metrics and generate recommendations"""
        try:
            self.logger.info("Starting performance analysis")
            
            # Get performance stats from all services
            smart_engine_stats = self.smart_engine.get_performance_stats()
            template_stats = self.template_service.get_performance_stats()
            ocr_stats = self.ocr_service.get_performance_stats()
            
            analysis_results = {
                'timestamp': time.time(),
                'overall_health': 'healthy',
                'performance_score': 0.0,
                'bottlenecks': [],
                'recommendations': [],
                'service_stats': {
                    'smart_engine': smart_engine_stats,
                    'template_service': template_stats,
                    'ocr_service': ocr_stats
                }
            }
            
            # Analyze overall success rates
            overall_success_rate = smart_engine_stats.get('overall_success_rate', 0)
            if overall_success_rate < 0.7:
                analysis_results['overall_health'] = 'warning'
                analysis_results['bottlenecks'].append('Low overall success rate')
                analysis_results['recommendations'].append('Review and update templates and OCR settings')
            elif overall_success_rate < 0.5:
                analysis_results['overall_health'] = 'critical'
            
            # Analyze response times
            avg_times = smart_engine_stats.get('average_execution_times', {})
            for method, avg_time in avg_times.items():
                if avg_time > 5.0:  # More than 5 seconds
                    analysis_results['bottlenecks'].append(f'Slow {method} method ({avg_time:.2f}s)')
                    analysis_results['recommendations'].append(f'Optimize {method} method performance')
            
            # Analyze error patterns
            error_categories = smart_engine_stats.get('error_categories', {})
            total_errors = sum(error_categories.values())
            if total_errors > 0:
                most_common_error = max(error_categories.items(), key=lambda x: x[1])
                if most_common_error[1] > total_errors * 0.3:  # More than 30% of errors
                    analysis_results['bottlenecks'].append(f'High {most_common_error[0]} errors')
                    analysis_results['recommendations'].append(f'Address {most_common_error[0]} issues')
            
            # Calculate performance score (0-100)
            score_factors = [
                overall_success_rate * 40,  # 40% weight
                min(1.0, 5.0 / max(avg_times.values(), default=1.0)) * 30,  # 30% weight for speed
                max(0, 1.0 - (total_errors / max(smart_engine_stats.get('total_operations', 1), 1))) * 30  # 30% weight for reliability
            ]
            analysis_results['performance_score'] = sum(score_factors)
            
            # Set health based on performance score
            if analysis_results['performance_score'] < 50:
                analysis_results['overall_health'] = 'critical'
            elif analysis_results['performance_score'] < 70:
                analysis_results['overall_health'] = 'warning'
            
            self.logger.info(f"Performance analysis completed: {analysis_results['overall_health']} (score: {analysis_results['performance_score']:.1f})")
            
            return analysis_results
            
        except Exception as e:
            self.logger.error(f"Performance analysis failed: {e}")
            return {'error': str(e)}    

    def analyze_errors(self) -> Dict[str, Any]:
        """Analyze error patterns and generate reports"""
        try:
            self.logger.info("Starting error analysis")
            
            # Get error data from all services
            smart_engine_errors = self.smart_engine.get_error_stats()
            template_errors = self.template_service.get_error_stats()
            ocr_errors = sel