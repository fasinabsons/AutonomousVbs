#!/usr/bin/env python3
"""
Sequential VBS Workflow: Phase 1 → Phase 2 → Phase 3
Complete automation workflow with proper sequencing and error handling
"""

import time
import logging
from datetime import datetime
from pathlib import Path
import sys

# Add project paths
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "vbs"))

# Import all phases
from vbs.vbs_phase1_login import VBSPhase1_Login
from vbs.vbs_phase2_navigation_fixed import VBSPhase2Clean
from vbs.vbs_phase3_upload_complete import VBSPhase3Complete
from vbs.vbs_phase4_report_fixed import VBSPhase4_ReportFixed

class SequentialVBSWorkflow:
    """Sequential execution of VBS workflow phases"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        self.phase1 = None
        self.phase2 = None
        self.phase3 = None
        self.results = {}
        
        self.logger.info("🚀 Sequential VBS Workflow initialized")
    
    def _setup_logging(self):
        """Setup logging for sequential workflow"""
        logger = logging.getLogger("SequentialVBSWorkflow")
        logger.setLevel(logging.INFO)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # Console handler
        console_handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - WORKFLOW - %(levelname)s - %(message)s')
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # File handler
        try:
            today_folder = datetime.now().strftime("%d%b").lower()
            log_dir = project_root / "EHC_Logs" / today_folder
            log_dir.mkdir(parents=True, exist_ok=True)
            
            log_file = log_dir / f"sequential_workflow_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            
            self.logger.info(f"📁 Workflow log: {log_file}")
            
        except Exception as e:
            logger.warning(f"Could not create log file: {e}")
        
        return logger
    
    def execute_complete_workflow(self):
        """Execute the complete workflow: Phase 1 → Phase 2 → Phase 3"""
        try:
            self.logger.info("🎯 Starting COMPLETE VBS WORKFLOW")
            self.logger.info("=" * 80)
            self.logger.info("WORKFLOW: Phase 1 (Login) → Phase 2 (Navigation) → Phase 3 (Upload)")
            self.logger.info("=" * 80)
            
            workflow_start_time = time.time()
            
            # ===== PHASE 1: LOGIN =====
            self.logger.info("🔑 PHASE 1: VBS LOGIN")
            self.logger.info("-" * 40)
            
            try:
                self.phase1 = VBSPhase1_Login()
                phase1_result = self.phase1.execute_vbs_login()
                self.results["phase1"] = phase1_result
                
                if not phase1_result.get("success", False):
                    self.logger.error("❌ PHASE 1 FAILED - Cannot proceed")
                    return self._generate_workflow_result(False, "Phase 1 login failed")
                
                self.logger.info("✅ PHASE 1 COMPLETED - VBS logged in successfully")
                
                # Wait for VBS to stabilize
                self.logger.info("⏳ Waiting 3 seconds for VBS to stabilize...")
                time.sleep(3.0)
                
            except Exception as e:
                self.logger.error(f"❌ PHASE 1 EXCEPTION: {e}")
                return self._generate_workflow_result(False, f"Phase 1 exception: {e}")
            
            # ===== PHASE 2: NAVIGATION =====
            self.logger.info("🧭 PHASE 2: VBS NAVIGATION")
            self.logger.info("-" * 40)
            
            try:
                self.phase2 = VBSPhase2Clean()
                phase2_result = self.phase2.execute_phase2()
                self.results["phase2"] = phase2_result
                
                if not phase2_result.get("success", False):
                    self.logger.error("❌ PHASE 2 FAILED - Cannot proceed to upload")
                    return self._generate_workflow_result(False, "Phase 2 navigation failed")
                
                self.logger.info("✅ PHASE 2 COMPLETED - Navigation successful")
                
                # Wait for form to be ready
                self.logger.info("⏳ Waiting 2 seconds for form to be ready...")
                time.sleep(2.0)
                
            except Exception as e:
                self.logger.error(f"❌ PHASE 2 EXCEPTION: {e}")
                return self._generate_workflow_result(False, f"Phase 2 exception: {e}")
            
            # ===== PHASE 3: UPLOAD =====
            self.logger.info("📤 PHASE 3: DATA UPLOAD")
            self.logger.info("-" * 40)
            
            try:
                self.phase3 = VBSPhase3Complete()
                phase3_result = self.phase3.execute_complete_phase3()
                self.results["phase3"] = phase3_result
                
                if not phase3_result.get("success", False):
                    self.logger.error("❌ PHASE 3 FAILED - Upload process failed")
                    return self._generate_workflow_result(False, "Phase 3 upload failed")
                
                self.logger.info("✅ PHASE 3 COMPLETED - Upload successful")
                
            except Exception as e:
                self.logger.error(f"❌ PHASE 3 EXCEPTION: {e}")
                return self._generate_workflow_result(False, f"Phase 3 exception: {e}")
            
            # ===== VBS RESTART FOR PHASE 4 =====
            self.logger.info("🔄 VBS RESTART FOR PHASE 4")
            self.logger.info("-" * 40)
            
            try:
                self.logger.info("⏳ Waiting 5 seconds before VBS restart...")
                time.sleep(5.0)
                
                # Restart VBS using Phase 1
                self.logger.info("🔑 Restarting VBS (Phase 1 for Phase 4)...")
                phase1_restart = VBSPhase1_Login()
                phase1_restart_result = phase1_restart.execute_vbs_login()
                
                if not phase1_restart_result.get("success", False):
                    self.logger.error("❌ VBS RESTART FAILED - Cannot run Phase 4")
                    return self._generate_workflow_result(False, "VBS restart for Phase 4 failed")
                
                self.logger.info("✅ VBS RESTARTED - Ready for Phase 4")
                time.sleep(3.0)
                
            except Exception as e:
                self.logger.error(f"❌ VBS RESTART EXCEPTION: {e}")
                return self._generate_workflow_result(False, f"VBS restart exception: {e}")
            
            # ===== PHASE 4: PDF REPORT =====
            self.logger.info("📄 PHASE 4: PDF REPORT GENERATION")
            self.logger.info("-" * 40)
            
            try:
                self.phase4 = VBSPhase4_ReportFixed()
                phase4_result = self.phase4.execute_report_generation()
                self.results["phase4"] = phase4_result
                
                if not phase4_result.get("success", False):
                    self.logger.error("❌ PHASE 4 FAILED - PDF generation failed")
                    return self._generate_workflow_result(False, "Phase 4 PDF generation failed")
                
                self.logger.info("✅ PHASE 4 COMPLETED - PDF report generated")
                
            except Exception as e:
                self.logger.error(f"❌ PHASE 4 EXCEPTION: {e}")
                return self._generate_workflow_result(False, f"Phase 4 exception: {e}")
            
            # ===== WORKFLOW COMPLETION =====
            workflow_time = time.time() - workflow_start_time
            hours = int(workflow_time / 3600)
            minutes = int((workflow_time % 3600) / 60)
            
            self.logger.info("🎉 COMPLETE WORKFLOW SUCCESSFUL!")
            self.logger.info("=" * 80)
            self.logger.info(f"⏱️ Total execution time: {hours}h {minutes}m")
            self.logger.info("✅ Phase 1: Login completed")
            self.logger.info("✅ Phase 2: Navigation completed")
            self.logger.info("✅ Phase 3: Upload completed")
            self.logger.info("✅ Phase 4: PDF report generated")
            self.logger.info("🔚 VBS application closed")
            self.logger.info("📧 Ready for email delivery")
            self.logger.info("=" * 80)
            
            return self._generate_workflow_result(True, "Complete workflow successful")
            
        except Exception as e:
            self.logger.error(f"❌ WORKFLOW CRITICAL ERROR: {e}")
            import traceback
            traceback.print_exc()
            return self._generate_workflow_result(False, f"Critical workflow error: {e}")
    
    def _generate_workflow_result(self, success: bool, message: str):
        """Generate comprehensive workflow result"""
        return {
            "success": success,
            "message": message,
            "timestamp": datetime.now().isoformat(),
            "phases": {
                "phase1": self.results.get("phase1", {"success": False, "error": "Not executed"}),
                "phase2": self.results.get("phase2", {"success": False, "error": "Not executed"}),
                "phase3": self.results.get("phase3", {"success": False, "error": "Not executed"}),
                "phase4": self.results.get("phase4", {"success": False, "error": "Not executed"})
            },
            "summary": {
                "phase1_success": self.results.get("phase1", {}).get("success", False),
                "phase2_success": self.results.get("phase2", {}).get("success", False),
                "phase3_success": self.results.get("phase3", {}).get("success", False),
                "phase4_success": self.results.get("phase4", {}).get("success", False),
                "upload_completed": self.results.get("phase3", {}).get("upload_success", False),
                "pdf_generated": self.results.get("phase4", {}).get("success", False)
            }
        }

def main():
    """Main execution function"""
    print("🚀 SEQUENTIAL VBS WORKFLOW")
    print("=" * 60)
    print("SEQUENCE: Phase 1 (Login) → Phase 2 (Navigation) → Phase 3 (Upload) → Phase 4 (PDF)")
    print("FEATURES:")
    print("✅ Automatic VBS login")
    print("✅ Enhanced navigation with keyboard fallbacks")
    print("✅ Reliable import button clicking")
    print("✅ 30 minutes to 3 hours upload monitoring")
    print("✅ Audio detection for popups")
    print("✅ Automatic VBS restart between phases")
    print("✅ PDF report generation")
    print("✅ Ready for email delivery")
    print("=" * 60)
    
    try:
        workflow = SequentialVBSWorkflow()
        result = workflow.execute_complete_workflow()
        
        print(f"\n📊 WORKFLOW RESULTS:")
        print(f"Success: {result['success']}")
        print(f"Message: {result['message']}")
        
        print(f"\n📋 PHASE SUMMARY:")
        summary = result['summary']
        print(f"   Phase 1 (Login): {'✅' if summary['phase1_success'] else '❌'}")
        print(f"   Phase 2 (Navigation): {'✅' if summary['phase2_success'] else '❌'}")
        print(f"   Phase 3 (Upload): {'✅' if summary['phase3_success'] else '❌'}")
        print(f"   Phase 4 (PDF Report): {'✅' if summary['phase4_success'] else '❌'}")
        print(f"   Upload Completed: {'✅' if summary['upload_completed'] else '❌'}")
        print(f"   PDF Generated: {'✅' if summary['pdf_generated'] else '❌'}")
        
        if result["success"]:
            print("\n🎉 COMPLETE WORKFLOW SUCCESSFUL!")
            print("📧 Ready for email delivery to General Manager")
        else:
            print(f"\n❌ WORKFLOW FAILED")
            
            # Show detailed errors
            for phase_name, phase_result in result['phases'].items():
                if not phase_result.get("success", False):
                    errors = phase_result.get("errors", [phase_result.get("error", "Unknown error")])
                    print(f"   {phase_name}: {errors}")
        
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("Sequential VBS Workflow Completed")

if __name__ == "__main__":
    main() 