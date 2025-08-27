#!/usr/bin/env python3
"""
VBS Automation Main Entry Point
Provides a simple interface to run VBS automation phases
"""

import sys
import os
import argparse
from typing import Optional, Dict, Any
import json

# Add the parent directory to the path so we can import our modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from vbs_core import VBSAutomator, run_complete_vbs_automation, run_vbs_phase
try:
    from utils.config_manager import ConfigManager
    from utils.file_manager import FileManager
except ImportError:
    # Fallback if utils not available
    ConfigManager = None
    FileManager = None

def main():
    """Main entry point for VBS automation"""
    parser = argparse.ArgumentParser(description='VBS Automation System')
    parser.add_argument('--phase', type=str, choices=['login', 'navigation', 'upload', 'report', 'all'], 
                       default='all', help='Phase to run (default: all)')
    parser.add_argument('--excel-file', type=str, help='Path to Excel file for upload phase')
    parser.add_argument('--config', type=str, help='Path to configuration file')
    parser.add_argument('--verbose', '-v', action='store_true', help='Enable verbose logging')
    parser.add_argument('--dry-run', action='store_true', help='Show what would be done without executing')
    
    args = parser.parse_args()
    
    print("VBS Automation System")
    print("=" * 50)
    
    # Initialize managers
    config_manager = None
    file_manager = None
    
    try:
        # Load configuration if provided
        if ConfigManager and args.config and os.path.exists(args.config):
            config_manager = ConfigManager(args.config)
            print(f"Loaded configuration from: {args.config}")
        elif ConfigManager:
            config_manager = ConfigManager()
            print("Using default configuration")
        else:
            config_manager = None
            print("Configuration manager not available")
        
        # Initialize file manager
        if FileManager:
            file_manager = FileManager()
        else:
            file_manager = None
            print("File manager not available")
        
        if args.dry_run:
            print("DRY RUN MODE - No actual execution")
            print(f"   Phase to run: {args.phase}")
            if args.excel_file:
                print(f"   Excel file: {args.excel_file}")
            return
        
        # Run the requested phase(s)
        if args.phase == 'all':
            print("Running complete VBS automation workflow...")
            result = run_complete_vbs_automation(
                excel_file_path=args.excel_file,
                config_manager=config_manager,
                file_manager=file_manager
            )
        else:
            print(f"Running VBS phase: {args.phase}")
            kwargs = {}
            if args.excel_file:
                kwargs['excel_file_path'] = args.excel_file
            
            result = run_vbs_phase(
                phase_name=args.phase,
                config_manager=config_manager,
                file_manager=file_manager,
                **kwargs
            )
        
        # Display results
        print("\n" + "=" * 50)
        print("EXECUTION RESULTS")
        print("=" * 50)
        
        if result['success']:
            print("SUCCESS")
            if 'phases_completed' in result:
                print(f"   Phases completed: {', '.join(result['phases_completed'])}")
            if 'total_duration_minutes' in result:
                print(f"   Total duration: {result['total_duration_minutes']} minutes")
        else:
            print("FAILED")
            if 'phases_failed' in result:
                print(f"   Phases failed: {', '.join(result['phases_failed'])}")
        
        if result.get('errors'):
            print("\nErrors:")
            for error in result['errors']:
                print(f"   - {error}")
        
        # Save results to file
        results_file = f"EHC_Logs/vbs_automation_results_{args.phase}.json"
        os.makedirs(os.path.dirname(results_file), exist_ok=True)
        
        with open(results_file, 'w') as f:
            json.dump(result, f, indent=2, default=str)
        
        print(f"\nResults saved to: {results_file}")
        
        # Exit with appropriate code
        sys.exit(0 if result['success'] else 1)
        
    except KeyboardInterrupt:
        print("\nInterrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

def run_login():
    """Convenience function to run just the login phase"""
    return run_vbs_phase('login')

def run_navigation():
    """Convenience function to run just the navigation phase"""
    return run_vbs_phase('navigation')

def run_upload(excel_file_path: Optional[str] = None):
    """Convenience function to run just the upload phase"""
    return run_vbs_phase('upload', excel_file_path=excel_file_path)

def run_report():
    """Convenience function to run just the report phase"""
    return run_vbs_phase('report')

def run_full_automation(excel_file_path: Optional[str] = None):
    """Convenience function to run the complete automation"""
    return run_complete_vbs_automation(excel_file_path)

# Example usage functions
def example_usage():
    """Show example usage"""
    print("""
VBS Automation System - Example Usage:

1. Run complete automation:
   python vbs_automator.py --phase all

2. Run specific phase:
   python vbs_automator.py --phase login
   python vbs_automator.py --phase navigation
   python vbs_automator.py --phase upload --excel-file path/to/file.xls
   python vbs_automator.py --phase report

3. Use custom configuration:
   python vbs_automator.py --config config/custom_settings.json

4. Dry run (show what would be done):
   python vbs_automator.py --phase all --dry-run

5. Verbose output:
   python vbs_automator.py --phase all --verbose

Programmatic usage:
   from vbs.vbs_automator import run_full_automation
   result = run_full_automation('path/to/excel/file.xls')
   print(f"Success: {result['success']}")
""")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        # No arguments provided, show help
        example_usage()
    else:
        main()