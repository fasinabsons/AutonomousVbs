#!/usr/bin/env python3
"""
Launcher script for MoonFlower WiFi Automation Configuration Panel
Provides easy access to the configuration interface
"""

import sys
import os
import logging
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

def main():
    """Main launcher function"""
    try:
        print("Starting MoonFlower WiFi Automation Configuration Panel...")
        
        # Setup basic logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Import and run configuration panel
        from gui.config_panel import ConfigurationPanel
        
        config_panel = ConfigurationPanel()
        config_panel.run()
        
    except ImportError as e:
        print(f"Import error: {e}")
        print("Please ensure all required dependencies are installed:")
        print("- tkinter (usually included with Python)")
        print("- Required project modules")
        sys.exit(1)
    
    except Exception as e:
        print(f"Failed to start configuration panel: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()