#!/usr/bin/env python3
"""
Simple script to run VBS Phase 4 directly
For testing Phase 4 after Phase 3 completes
"""

import sys
import subprocess
from pathlib import Path

def main():
    """Run VBS Phase 4 directly"""
    print("🚀 Running VBS Phase 4 directly...")
    print("=" * 50)
    
    # Change to project root
    project_root = Path(__file__).parent
    
    try:
        # Run Phase 4
        result = subprocess.run([
            sys.executable, 
            str(project_root / "vbs" / "vbs_phase4_report_fixed.py")
        ], capture_output=False, text=True)
        
        if result.returncode == 0:
            print("✅ VBS Phase 4 completed successfully!")
            return 0
        else:
            print(f"❌ VBS Phase 4 failed with exit code: {result.returncode}")
            return 1
            
    except Exception as e:
        print(f"❌ Error running VBS Phase 4: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())