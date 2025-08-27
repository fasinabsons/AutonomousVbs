#!/usr/bin/env python3
"""
VBS Automation Core Module
Provides shared functionality and orchestration for all VBS automation phases
"""

import os
import time
import logging
from datetime import datetime
from typing import Dict, Optional, Any
import traceback

# Import all VBS phases
from vbs_phase1_login import VBSPhase1_Login
from vbs_phase2_navigation_fixed import VBSPhase2_NavigationFixed
from vbs_phase3_upload_fixed import VBSPhase3_UploadFixed
from vbs_phase4_report_fixed import VBSPhase4_ReportFixed

class VBSAutomator:
    """Main VBS Automation orchestrator that manages all phases"""
    
    def __init__(self, config_manager=None, file_manager=None):
        """Initialize VBS Automator with configuration"""
        self.logger = self._setup_logging()
        self.config_manager = config_manager
        self.file_manager = file_manager
        
        # Phase instances
        self.login_phase = None
        self.navigation_phase = None
        self.upload_phase = None
        self.report_phase = None
        
        # Execution state
        self.current_window_handle = None
        self.current_process_id = None
        self.execution_results = {}
        
        self.logger.info("VBS Automator initialized")
    
    def _setup_logging(self) -> logging.Logger:
        """Setup enhanced logging"""
        logger = logging.getLogger("VBSAutomator")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Add file handler
            try:
                log_dir = "EHC_Logs"
                os.makedirs(log_dir, exist_ok=True)
                log_file = os.path.join(log_dir, f"vbs_automator_{datetime.now().strftime('%Y%m%d')}.log")
                file_handler = logging.FileHandler(log_file)
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except Exception as e:
                print(f"Warning: Could not set up file logging: {e}")
        
        return logger
    
    def run_complete_workflow(self, excel_file_path: Optional[str] = None) -> Dict[str, Any]:
        """Run the complete VBS automation workflow"""
        try:
            self.logger.info("Starting Complete VBS Automation Workflow")
            
            workflow_result = {
                "success": False,
                "start_time": datetime.now().isoformat(),
                "phases_completed": [],
                "phases_failed": [],
                "errors": [],
                "total_duration_minutes": 0
            }
            
            start_time = time.time()
            
            # Phase 1: Login
            self.logger.info("=" * 60)
            self.logger.info("PHASE 1: VBS LOGIN")
            self.logger.info("=" * 60)
            
            login_result = self.run_login_phase()
            if login_result["success"]:
                workflow_result["phases_completed"].append("Login")
                self.current_window_handle = login_result.get("window_handle")
                self.current_process_id = login_result.get("process_id")
                self.execution_results["login"] = login_result
            else:
                workflow_result["phases_failed"].append("Login")
                workflow_result["errors"].extend(login_result.get("errors", []))
                return workflow_result
            
            # Phase 2: Navigation
            self.logger.info("=" * 60)
            self.logger.info("PHASE 2: VBS NAVIGATION")
            self.logger.info("=" * 60)
            
            navigation_result = self.run_navigation_phase()
            if navigation_result["success"]:
                workflow_result["phases_completed"].append("Navigation")
                self.current_window_handle = navigation_result.get("window_handle", self.current_window_handle)
                self.execution_results["navigation"] = navigation_result
            else:
                workflow_result["phases_failed"].append("Navigation")
                workflow_result["errors"].extend(navigation_result.get("errors", []))
                return workflow_result
            
            # Phase 3: Upload
            self.logger.info("=" * 60)
            self.logger.info("PHASE 3: DATA UPLOAD")
            self.logger.info("=" * 60)
            
            upload_result = self.run_upload_phase(excel_file_path)
            if upload_result["success"]:
                workflow_result["phases_completed"].append("Upload")
                self.execution_results["upload"] = upload_result
                
                # VBS needs to be restarted after upload
                if upload_result.get("requires_restart"):
                    self.logger.info("VBS restart required after upload")
                    time.sleep(5)  # Wait for application to close
                    
                    # Re-login for Phase 4
                    self.logger.info("Re-logging in for Phase 4...")
                    login_result = self.run_login_phase()
                    if not login_result["success"]:
                        workflow_result["phases_failed"].append("Upload-Restart-Login")
                        workflow_result["errors"].extend(login_result.get("errors", []))
                        return workflow_result
                    
                    self.current_window_handle = login_result.get("window_handle")
                    self.current_process_id = login_result.get("process_id")
            else:
                workflow_result["phases_failed"].append("Upload")
                workflow_result["errors"].extend(upload_result.get("errors", []))
                return workflow_result
            
            # Phase 4: Report Generation
            self.logger.info("=" * 60)
            self.logger.info("PHASE 4: PDF REPORT GENERATION")
            self.logger.info("=" * 60)
            
            report_result = self.run_report_phase()
            if report_result["success"]:
                workflow_result["phases_completed"].append("Report")
                self.execution_results["report"] = report_result
            else:
                workflow_result["phases_failed"].append("Report")
                workflow_result["errors"].extend(report_result.get("errors", []))
                return workflow_result
            
            # Calculate total duration
            total_duration = (time.time() - start_time) / 60
            workflow_result["total_duration_minutes"] = int(total_duration)
            
            # Success
            workflow_result.update({
                "success": True,
                "end_time": datetime.now().isoformat(),
                "execution_results": self.execution_results
            })
            
            self.logger.info("=" * 60)
            self.logger.info("COMPLETE VBS AUTOMATION WORKFLOW SUCCESSFUL!")
            self.logger.info(f"Total Duration: {total_duration:.1f} minutes")
            self.logger.info(f"Phases Completed: {', '.join(workflow_result['phases_completed'])}")
            self.logger.info("=" * 60)
            
            return workflow_result
            
        except Exception as e:
            error_msg = f"Complete workflow failed: {str(e)}"
            self.logger.error(f"ERROR: {error_msg}")
            self.logger.error(traceback.format_exc())
            workflow_result["errors"].append(error_msg)
            return workflow_result
    
    def run_login_phase(self) -> Dict[str, Any]:
        """Run Phase 1: Login"""
        try:
            self.login_phase = VBSPhase1_Login()
            result = self.login_phase.execute_vbs_login()
            
            if result["success"]:
                self.logger.info("SUCCESS: Phase 1 (Login) completed successfully")
            else:
                self.logger.error("FAILED: Phase 1 (Login) failed")
            
            return result
            
        except Exception as e:
            self.logger.error(f"ERROR: Phase 1 (Login) exception: {e}")
            return {"success": False, "phase": "Login", "errors": [str(e)]}
    
    def run_navigation_phase(self) -> Dict[str, Any]:
        """Run Phase 2: Navigation"""
        try:
            self.navigation_phase = VBSPhase2_NavigationFixed()
            result = self.navigation_phase.execute_navigation()
            
            if result["success"]:
                self.logger.info("SUCCESS: Phase 2 (Navigation) completed successfully")
            else:
                self.logger.error("FAILED: Phase 2 (Navigation) failed")
            
            return result
            
        except Exception as e:
            self.logger.error(f"ERROR: Phase 2 (Navigation) exception: {e}")
            return {"success": False, "phase": "Navigation", "errors": [str(e)]}
    
    def run_upload_phase(self, excel_file_path: Optional[str] = None) -> Dict[str, Any]:
        """Run Phase 3: Upload"""
        try:
            self.upload_phase = VBSPhase3_UploadFixed()
            result = self.upload_phase.execute_upload_process()
            
            if result["success"]:
                self.logger.info("SUCCESS: Phase 3 (Upload) completed successfully")
            else:
                self.logger.error("FAILED: Phase 3 (Upload) failed")
            
            return result
            
        except Exception as e:
            self.logger.error(f"ERROR: Phase 3 (Upload) exception: {e}")
            return {"success": False, "phase": "Upload", "errors": [str(e)]}
    
    def run_report_phase(self) -> Dict[str, Any]:
        """Run Phase 4: Report"""
        try:
            self.report_phase = VBSPhase4_ReportFixed()
            result = self.report_phase.execute_report_generation()
            
            if result["success"]:
                self.logger.info("SUCCESS: Phase 4 (Report) completed successfully")
            else:
                self.logger.error("FAILED: Phase 4 (Report) failed")
            
            return result
            
        except Exception as e:
            self.logger.error(f"ERROR: Phase 4 (Report) exception: {e}")
            return {"success": False, "phase": "Report", "errors": [str(e)]}
    
    def run_individual_phase(self, phase_name: str, **kwargs) -> Dict[str, Any]:
        """Run an individual phase by name"""
        try:
            phase_name_lower = phase_name.lower()
            
            if phase_name_lower == "login":
                return self.run_login_phase()
            elif phase_name_lower == "navigation":
                return self.run_navigation_phase()
            elif phase_name_lower == "upload":
                excel_file_path = kwargs.get("excel_file_path")
                return self.run_upload_phase(excel_file_path)
            elif phase_name_lower == "report":
                return self.run_report_phase()
            else:
                return {
                    "success": False,
                    "phase": phase_name,
                    "errors": [f"Unknown phase: {phase_name}"]
                }
                
        except Exception as e:
            self.logger.error(f"ERROR: Individual phase {phase_name} exception: {e}")
            return {"success": False, "phase": phase_name, "errors": [str(e)]}
    
    def get_execution_summary(self) -> Dict[str, Any]:
        """Get summary of execution results"""
        try:
            summary = {
                "total_phases": len(self.execution_results),
                "successful_phases": [],
                "failed_phases": [],
                "phase_details": {}
            }
            
            for phase_name, result in self.execution_results.items():
                if result.get("success", False):
                    summary["successful_phases"].append(phase_name)
                else:
                    summary["failed_phases"].append(phase_name)
                
                summary["phase_details"][phase_name] = {
                    "success": result.get("success", False),
                    "duration": result.get("total_duration_minutes", 0),
                    "steps_completed": len(result.get("steps_completed", [])),
                    "errors": result.get("errors", [])
                }
            
            return summary
            
        except Exception as e:
            self.logger.error(f"ERROR: Execution summary failed: {e}")
            return {"error": str(e)}
    
    def cleanup(self):
        """Clean up resources"""
        try:
            self.logger.info("Cleaning up VBS Automator resources...")
            
            # Close VBS application if still running
            if self.login_phase:
                self.login_phase.close_application()
            
            self.logger.info("SUCCESS: VBS Automator cleanup completed")
            
        except Exception as e:
            self.logger.error(f"ERROR: Cleanup failed: {e}")

# Convenience functions for direct usage
def run_complete_vbs_automation(excel_file_path: Optional[str] = None, config_manager=None, file_manager=None) -> Dict[str, Any]:
    """Run the complete VBS automation workflow"""
    automator = VBSAutomator(config_manager, file_manager)
    try:
        return automator.run_complete_workflow(excel_file_path)
    finally:
        automator.cleanup()

def run_vbs_phase(phase_name: str, config_manager=None, file_manager=None, **kwargs) -> Dict[str, Any]:
    """Run a single VBS phase"""
    automator = VBSAutomator(config_manager, file_manager)
    try:
        return automator.run_individual_phase(phase_name, **kwargs)
    finally:
        automator.cleanup()

if __name__ == "__main__":
    # Test the complete workflow
    print("Testing Complete VBS Automation Workflow")
    print("=" * 80)
    
    result = run_complete_vbs_automation()
    
    print(f"\nWorkflow Results:")
    print(f"   Success: {result['success']}")
    print(f"   Phases Completed: {result.get('phases_completed', [])}")
    print(f"   Phases Failed: {result.get('phases_failed', [])}")
    print(f"   Total Duration: {result.get('total_duration_minutes', 0)} minutes")
    
    if result.get('errors'):
        print(f"\nErrors:")
        for error in result['errors']:
            print(f"   - {error}")
    
    if result["success"]:
        print("\nSUCCESS: Complete VBS automation workflow completed successfully!")
    else:
        print(f"\nFAILED: VBS automation workflow failed")
    
    print("\n" + "=" * 80)
    print("VBS Automation Test Completed")