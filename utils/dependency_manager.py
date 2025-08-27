#!/usr/bin/env python3
"""
Dependency Manager for MoonFlower CSV Download System
Automatically checks and installs required Python dependencies
"""

import sys
import subprocess
import logging
import importlib
from pathlib import Path
from typing import Dict, List, Tuple, Any
import pkg_resources
from datetime import datetime

class DependencyManager:
    """Manages Python dependencies with automatic installation"""
    
    def __init__(self, requirements_file: str = "requirements.txt"):
        self.requirements_file = Path(requirements_file)
        self.logger = self._setup_logging()
        
        # Required modules for the system
        self.required_modules = {
            'selenium': '4.15.2',
            'undetected-chromedriver': '3.5.4', 
            'beautifulsoup4': '4.12.2',
            'requests': '2.31.0',
            'urllib3': '2.0.7',
            'pyautogui': '0.9.54',
            'opencv-python': '4.8.1.78',
            'pytesseract': '0.3.10',
            'pillow': '10.0.1',
            'pandas': '2.1.4',
            'xlwt': '1.3.0',
            'pywin32': '306',
            'schedule': '1.2.0',
            'pathlib': None,  # Built-in module
            'typing': None    # Built-in module
        }
        
        self.installation_results = {}
        
    def _setup_logging(self) -> logging.Logger:
        """Setup logging for dependency manager"""
        logger = logging.getLogger("DependencyManager")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        return logger
    
    def check_dependencies(self) -> Dict[str, bool]:
        """Check if all required dependencies are installed"""
        self.logger.info("Checking Python dependencies...")
        
        dependency_status = {}
        missing_modules = []
        
        for module_name, required_version in self.required_modules.items():
            try:
                # Try to import the module
                if module_name == 'opencv-python':
                    # opencv-python imports as cv2
                    importlib.import_module('cv2')
                elif module_name == 'pillow':
                    # pillow imports as PIL
                    importlib.import_module('PIL')
                elif module_name == 'beautifulsoup4':
                    # beautifulsoup4 imports as bs4
                    importlib.import_module('bs4')
                elif module_name == 'undetected-chromedriver':
                    # undetected-chromedriver imports as undetected_chromedriver
                    importlib.import_module('undetected_chromedriver')
                else:
                    importlib.import_module(module_name)
                
                dependency_status[module_name] = True
                self.logger.info(f"✓ {module_name} - Available")
                
            except ImportError:
                dependency_status[module_name] = False
                missing_modules.append(module_name)
                self.logger.warning(f"✗ {module_name} - Missing")
        
        if missing_modules:
            self.logger.warning(f"Missing modules: {', '.join(missing_modules)}")
        else:
            self.logger.info("All dependencies are satisfied")
        
        return dependency_status
    
    def install_missing_dependencies(self) -> bool:
        """Attempt to install missing dependencies automatically"""
        self.logger.info("Starting automatic dependency installation...")
        
        dependency_status = self.check_dependencies()
        missing_modules = [module for module, status in dependency_status.items() if not status]
        
        if not missing_modules:
            self.logger.info("No missing dependencies to install")
            return True
        
        # Try to install from requirements.txt first
        if self.requirements_file.exists():
            self.logger.info(f"Installing from {self.requirements_file}")
            if self._install_from_requirements():
                # Re-check dependencies after installation
                new_status = self.check_dependencies()
                still_missing = [module for module, status in new_status.items() if not status]
                if not still_missing:
                    self.logger.info("All dependencies installed successfully from requirements.txt")
                    return True
        
        # Install individual missing modules
        success_count = 0
        for module_name in missing_modules:
            if module_name in ['pathlib', 'typing']:
                # Skip built-in modules
                continue
                
            if self._install_single_module(module_name):
                success_count += 1
        
        # Final dependency check
        final_status = self.check_dependencies()
        final_missing = [module for module, status in final_status.items() if not status]
        
        if not final_missing:
            self.logger.info("All dependencies installed successfully")
            return True
        else:
            self.logger.error(f"Failed to install: {', '.join(final_missing)}")
            return False
    
    def _install_from_requirements(self) -> bool:
        """Install dependencies from requirements.txt file"""
        try:
            self.logger.info("Installing packages from requirements.txt...")
            
            cmd = [sys.executable, "-m", "pip", "install", "-r", str(self.requirements_file)]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=300  # 5 minute timeout
            )
            
            if result.returncode == 0:
                self.logger.info("Requirements.txt installation completed successfully")
                return True
            else:
                self.logger.error(f"Requirements.txt installation failed: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error("Requirements.txt installation timed out")
            return False
        except Exception as e:
            self.logger.error(f"Requirements.txt installation error: {e}")
            return False
    
    def _install_single_module(self, module_name: str) -> bool:
        """Install a single Python module using pip"""
        try:
            required_version = self.required_modules.get(module_name)
            
            if required_version:
                package_spec = f"{module_name}=={required_version}"
            else:
                package_spec = module_name
            
            self.logger.info(f"Installing {package_spec}...")
            
            cmd = [sys.executable, "-m", "pip", "install", package_spec]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=180  # 3 minute timeout per module
            )
            
            if result.returncode == 0:
                self.logger.info(f"✓ {module_name} installed successfully")
                self.installation_results[module_name] = True
                return True
            else:
                self.logger.error(f"✗ {module_name} installation failed: {result.stderr}")
                self.installation_results[module_name] = False
                return False
                
        except subprocess.TimeoutExpired:
            self.logger.error(f"✗ {module_name} installation timed out")
            self.installation_results[module_name] = False
            return False
        except Exception as e:
            self.logger.error(f"✗ {module_name} installation error: {e}")
            self.installation_results[module_name] = False
            return False
    
    def generate_requirements_report(self) -> Dict[str, Any]:
        """Generate a detailed report of dependency status"""
        dependency_status = self.check_dependencies()
        
        report = {
            "timestamp": datetime.now().isoformat(),
            "python_version": sys.version,
            "total_dependencies": len(self.required_modules),
            "satisfied_dependencies": sum(dependency_status.values()),
            "missing_dependencies": len(self.required_modules) - sum(dependency_status.values()),
            "dependency_details": dependency_status,
            "installation_results": self.installation_results,
            "requirements_file_exists": self.requirements_file.exists()
        }
        
        return report
    
    def validate_critical_modules(self) -> Tuple[bool, List[str]]:
        """Validate that critical modules for CSV download are available"""
        critical_modules = ['selenium', 'beautifulsoup4', 'requests', 'pandas', 'xlwt']
        
        missing_critical = []
        
        for module in critical_modules:
            try:
                if module == 'beautifulsoup4':
                    importlib.import_module('bs4')
                else:
                    importlib.import_module(module)
            except ImportError:
                missing_critical.append(module)
        
        is_valid = len(missing_critical) == 0
        
        if is_valid:
            self.logger.info("All critical modules are available")
        else:
            self.logger.error(f"Missing critical modules: {', '.join(missing_critical)}")
        
        return is_valid, missing_critical
    
    def create_requirements_file(self) -> bool:
        """Create requirements.txt file if it doesn't exist"""
        if self.requirements_file.exists():
            self.logger.info("Requirements.txt already exists")
            return True
        
        try:
            requirements_content = []
            for module_name, version in self.required_modules.items():
                if version and module_name not in ['pathlib', 'typing']:
                    requirements_content.append(f"{module_name}=={version}")
            
            with open(self.requirements_file, 'w') as f:
                f.write('\n'.join(requirements_content))
            
            self.logger.info(f"Created {self.requirements_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to create requirements.txt: {e}")
            return False

def main():
    """Test the dependency manager"""
    manager = DependencyManager()
    
    print("=== Dependency Manager Test ===")
    
    # Check current status
    status = manager.check_dependencies()
    print(f"\nDependency Status: {sum(status.values())}/{len(status)} satisfied")
    
    # Try to install missing dependencies
    if not all(status.values()):
        print("\nAttempting to install missing dependencies...")
        success = manager.install_missing_dependencies()
        print(f"Installation result: {'Success' if success else 'Failed'}")
    
    # Generate report
    report = manager.generate_requirements_report()
    print(f"\nFinal Report:")
    print(f"- Total dependencies: {report['total_dependencies']}")
    print(f"- Satisfied: {report['satisfied_dependencies']}")
    print(f"- Missing: {report['missing_dependencies']}")
    
    # Validate critical modules
    is_valid, missing = manager.validate_critical_modules()
    print(f"\nCritical modules valid: {is_valid}")
    if missing:
        print(f"Missing critical modules: {', '.join(missing)}")

if __name__ == "__main__":
    main()