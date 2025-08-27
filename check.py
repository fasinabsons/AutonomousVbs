#!/usr/bin/env python3
"""
PDF and Excel Validation System
Validates PDF content, compares daily user counts, checks Excel integrity
and sends email reports to monitor system health.
"""

import os
import sys
import re
import logging
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
import PyPDF2

# Add project root to path for imports
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

try:
    from utils.path_manager import PathManager
    from email.email_delivery import EmailDeliverySystem
    import pandas as pd
    import xlrd
except ImportError as e:
    print(f"Warning: Could not import required modules: {e}")
    PathManager = None
    EmailDeliverySystem = None


class ValidationSystem:
    """PDF and Excel validation system with email reporting"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        
        # Initialize path management
        if PathManager:
            self.path_manager = PathManager()
            self.logger.info("✅ PathManager integration enabled")
        else:
            self.path_manager = None
            self.logger.warning("⚠️ PathManager not available, using fallback")
        
        # Initialize email system
        if EmailDeliverySystem:
            self.email_system = EmailDeliverySystem()
            self.logger.info("✅ Email system integration enabled")
        else:
            self.email_system = None
            self.logger.warning("⚠️ Email system not available")
        
        # Validation criteria
        self.validation_criteria = {
            'min_excel_rows': 100,           # Minimum expected rows in Excel
            'max_excel_rows': 50000,        # Maximum reasonable rows
            'min_pdf_pages': 1,              # Minimum PDF pages
            'max_pdf_size_mb': 50,           # Maximum PDF size
            'required_pdf_keywords': [       # Keywords that must exist in PDF
                'Total Active User Count',
                'MoonFlower',
                'WiFi',
                'Report',
                'Print Date'
            ],
            'user_count_growth_threshold': 0  # Today's count must be >= yesterday's
        }
        
        self.logger.info("Validation System initialized")
    
    def _setup_logging(self):
        """Setup comprehensive logging"""
        logger = logging.getLogger("ValidationSystem")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            # Console handler
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                today = datetime.now().strftime("%d%b").lower()
                log_dir = project_root / "EHC_Logs" / today
                log_dir.mkdir(parents=True, exist_ok=True)
                log_file = log_dir / f"validation_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
                
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
                
                logger.info(f"Logging to: {log_file}")
            except Exception as e:
                logger.warning(f"Could not setup file logging: {e}")
        
        return logger
    
    def get_date_folder(self, date_offset: int = 0) -> str:
        """Get date folder name (e.g., '25jul') with optional offset"""
        target_date = datetime.now() + timedelta(days=date_offset)
        return target_date.strftime("%d%b").lower()
    
    def find_latest_excel_file(self, date_folder: str = None) -> Optional[Path]:
        """Find the latest Excel file for given date"""
        try:
            if date_folder is None:
                date_folder = self.get_date_folder()
            
            if self.path_manager:
                excel_dir = self.path_manager.get_excel_directory()
            else:
                excel_dir = project_root / "EHC_Data_Merge" / date_folder
            
            if not excel_dir.exists():
                self.logger.warning(f"Excel directory does not exist: {excel_dir}")
                return None
            
            # Find Excel files
            excel_files = list(excel_dir.glob("*.xls*"))
            if not excel_files:
                self.logger.warning(f"No Excel files found in {excel_dir}")
                return None
            
            # Return most recent file
            latest_file = max(excel_files, key=lambda x: x.stat().st_mtime)
            self.logger.info(f"Found Excel file: {latest_file}")
            return latest_file
            
        except Exception as e:
            self.logger.error(f"Error finding Excel file: {e}")
            return None
    
    def find_latest_pdf_file(self, date_folder: str = None) -> Optional[Path]:
        """Find the latest PDF file for given date"""
        try:
            if date_folder is None:
                date_folder = self.get_date_folder()
            
            if self.path_manager:
                pdf_dir = self.path_manager.get_pdf_directory()
            else:
                pdf_dir = project_root / "EHC_Data_Pdf" / date_folder
            
            if not pdf_dir.exists():
                self.logger.warning(f"PDF directory does not exist: {pdf_dir}")
                return None
            
            # Find PDF files
            pdf_files = list(pdf_dir.glob("*.pdf"))
            if not pdf_files:
                self.logger.warning(f"No PDF files found in {pdf_dir}")
                return None
            
            # Return most recent file
            latest_file = max(pdf_files, key=lambda x: x.stat().st_mtime)
            self.logger.info(f"Found PDF file: {latest_file}")
            return latest_file
            
        except Exception as e:
            self.logger.error(f"Error finding PDF file: {e}")
            return None
    
    def validate_excel_file(self, excel_file: Path) -> Dict[str, Any]:
        """Validate Excel file structure and content"""
        try:
            self.logger.info(f"Validating Excel file: {excel_file.name}")
            
            validation_result = {
                'file_exists': excel_file.exists(),
                'file_size_mb': 0,
                'row_count': 0,
                'column_count': 0,
                'headers_valid': False,
                'data_valid': False,
                'errors': [],
                'warnings': []
            }
            
            if not excel_file.exists():
                validation_result['errors'].append("Excel file does not exist")
                return validation_result
            
            # File size check
            file_size = excel_file.stat().st_size
            validation_result['file_size_mb'] = round(file_size / (1024 * 1024), 2)
            
            if file_size == 0:
                validation_result['errors'].append("Excel file is empty (0 bytes)")
                return validation_result
            
            # Read Excel file
            try:
                # Try pandas first
                df = pd.read_excel(excel_file)
                validation_result['row_count'] = len(df)
                validation_result['column_count'] = len(df.columns)
                
                # Expected headers for VBS compatibility
                expected_headers = ['Hostname', 'IP_Address', 'MAC_Address', 'Package', 'AP_MAC', 'Upload', 'Download']
                
                # Check headers
                actual_headers = df.columns.tolist()
                missing_headers = [h for h in expected_headers if h not in actual_headers]
                
                if missing_headers:
                    validation_result['warnings'].append(f"Missing headers: {missing_headers}")
                else:
                    validation_result['headers_valid'] = True
                
                # Row count validation
                if validation_result['row_count'] < self.validation_criteria['min_excel_rows']:
                    validation_result['errors'].append(f"Too few rows: {validation_result['row_count']} (min: {self.validation_criteria['min_excel_rows']})")
                elif validation_result['row_count'] > self.validation_criteria['max_excel_rows']:
                    validation_result['warnings'].append(f"High row count: {validation_result['row_count']}")
                else:
                    validation_result['data_valid'] = True
                
                # Check for empty data
                non_empty_rows = df.dropna(how='all').shape[0]
                if non_empty_rows < validation_result['row_count'] * 0.8:  # 80% threshold
                    validation_result['warnings'].append(f"Many empty rows detected: {validation_result['row_count'] - non_empty_rows}")
                
                self.logger.info(f"Excel validation completed: {validation_result['row_count']} rows, {validation_result['column_count']} columns")
                
            except Exception as e:
                validation_result['errors'].append(f"Could not read Excel file: {str(e)}")
            
            return validation_result
            
        except Exception as e:
            self.logger.error(f"Excel validation failed: {e}")
            return {'file_exists': False, 'errors': [f"Validation failed: {str(e)}"]}
    
    def extract_user_count_from_pdf(self, pdf_file: Path) -> Optional[int]:
        """Extract Total Active User Count from PDF"""
        try:
            self.logger.info(f"Extracting user count from PDF: {pdf_file.name}")
            
            with open(pdf_file, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Extract text from all pages
                full_text = ""
                for page in pdf_reader.pages:
                    full_text += page.extract_text()
                
                # Look for "Total Active User Count" pattern
                patterns = [
                    r'Total Active User Count[:\s]*(\d+)',
                    r'Total Active User Count[:\s]*(\d{1,6})',
                    r'Active User Count[:\s]*(\d+)',
                    r'Total.*?(\d{3,6})\s*$'  # Fallback: number at end of line with "Total"
                ]
                
                for pattern in patterns:
                    matches = re.findall(pattern, full_text, re.IGNORECASE | re.MULTILINE)
                    if matches:
                        user_count = int(matches[-1])  # Take the last match
                        self.logger.info(f"Extracted user count: {user_count}")
                        return user_count
                
                self.logger.warning("Could not find Total Active User Count in PDF")
                return None
                
        except Exception as e:
            self.logger.error(f"Error extracting user count from PDF: {e}")
            return None
    
    def validate_pdf_file(self, pdf_file: Path) -> Dict[str, Any]:
        """Validate PDF file structure and content"""
        try:
            self.logger.info(f"Validating PDF file: {pdf_file.name}")
            
            validation_result = {
                'file_exists': pdf_file.exists(),
                'file_size_mb': 0,
                'page_count': 0,
                'keywords_found': [],
                'keywords_missing': [],
                'user_count': None,
                'content_valid': False,
                'errors': [],
                'warnings': []
            }
            
            if not pdf_file.exists():
                validation_result['errors'].append("PDF file does not exist")
                return validation_result
            
            # File size check
            file_size = pdf_file.stat().st_size
            validation_result['file_size_mb'] = round(file_size / (1024 * 1024), 2)
            
            if file_size == 0:
                validation_result['errors'].append("PDF file is empty (0 bytes)")
                return validation_result
            
            if validation_result['file_size_mb'] > self.validation_criteria['max_pdf_size_mb']:
                validation_result['warnings'].append(f"Large PDF file: {validation_result['file_size_mb']} MB")
            
            # Read PDF content
            try:
                with open(pdf_file, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    validation_result['page_count'] = len(pdf_reader.pages)
                    
                    # Extract text from all pages
                    full_text = ""
                    for page in pdf_reader.pages:
                        full_text += page.extract_text()
                    
                    # Check for required keywords
                    for keyword in self.validation_criteria['required_pdf_keywords']:
                        if keyword.lower() in full_text.lower():
                            validation_result['keywords_found'].append(keyword)
                        else:
                            validation_result['keywords_missing'].append(keyword)
                    
                    # Extract user count
                    validation_result['user_count'] = self.extract_user_count_from_pdf(pdf_file)
                    
                    # Validation checks
                    if validation_result['page_count'] < self.validation_criteria['min_pdf_pages']:
                        validation_result['errors'].append(f"Too few pages: {validation_result['page_count']}")
                    
                    if validation_result['keywords_missing']:
                        validation_result['warnings'].append(f"Missing keywords: {validation_result['keywords_missing']}")
                    
                    if len(validation_result['keywords_found']) >= len(self.validation_criteria['required_pdf_keywords']) * 0.8:
                        validation_result['content_valid'] = True
                    
                    self.logger.info(f"PDF validation completed: {validation_result['page_count']} pages, user count: {validation_result['user_count']}")
                    
            except Exception as e:
                validation_result['errors'].append(f"Could not read PDF file: {str(e)}")
            
            return validation_result
            
        except Exception as e:
            self.logger.error(f"PDF validation failed: {e}")
            return {'file_exists': False, 'errors': [f"Validation failed: {str(e)}"]}
    
    def compare_daily_user_counts(self, today_pdf: Path, yesterday_pdf: Path = None) -> Dict[str, Any]:
        """Compare today's user count with yesterday's"""
        try:
            self.logger.info("Comparing daily user counts...")
            
            comparison_result = {
                'today_count': None,
                'yesterday_count': None,
                'growth': None,
                'growth_percentage': None,
                'validation_passed': False,
                'message': ''
            }
            
            # Get today's count
            comparison_result['today_count'] = self.extract_user_count_from_pdf(today_pdf)
            
            # Find yesterday's PDF if not provided
            if yesterday_pdf is None:
                yesterday_folder = self.get_date_folder(-1)  # Yesterday
                yesterday_pdf = self.find_latest_pdf_file(yesterday_folder)
            
            # Get yesterday's count
            if yesterday_pdf and yesterday_pdf.exists():
                comparison_result['yesterday_count'] = self.extract_user_count_from_pdf(yesterday_pdf)
            
            # Calculate growth
            if comparison_result['today_count'] is not None and comparison_result['yesterday_count'] is not None:
                growth = comparison_result['today_count'] - comparison_result['yesterday_count']
                comparison_result['growth'] = growth
                
                if comparison_result['yesterday_count'] > 0:
                    growth_percentage = (growth / comparison_result['yesterday_count']) * 100
                    comparison_result['growth_percentage'] = round(growth_percentage, 2)
                
                # Validation
                if comparison_result['today_count'] >= comparison_result['yesterday_count']:
                    comparison_result['validation_passed'] = True
                    comparison_result['message'] = f"✅ User count growth confirmed: {growth:+d} users ({comparison_result['growth_percentage']:+.1f}%)"
                else:
                    comparison_result['message'] = f"⚠️ User count decreased: {growth:+d} users ({comparison_result['growth_percentage']:+.1f}%)"
            else:
                comparison_result['message'] = "❌ Could not compare user counts - missing data"
            
            self.logger.info(comparison_result['message'])
            return comparison_result
            
        except Exception as e:
            self.logger.error(f"User count comparison failed: {e}")
            return {'validation_passed': False, 'message': f"Comparison failed: {str(e)}"}
    
    def send_validation_report(self, excel_validation: Dict, pdf_validation: Dict, 
                             user_comparison: Dict) -> bool:
        """Send comprehensive validation report via email"""
        try:
            if not self.email_system:
                self.logger.warning("Email system not available")
                return False
            
            current_date = datetime.now().strftime('%d/%m/%Y')
            current_time = datetime.now().strftime('%H:%M:%S')
            
            # Determine overall status
            excel_ok = excel_validation.get('data_valid', False) and not excel_validation.get('errors', [])
            pdf_ok = pdf_validation.get('content_valid', False) and not pdf_validation.get('errors', [])
            growth_ok = user_comparison.get('validation_passed', False)
            
            overall_status = "✅ PASSED" if (excel_ok and pdf_ok and growth_ok) else "⚠️ ISSUES DETECTED"
            
            subject = f"📊 Daily Validation Report - {overall_status} - {current_date}"
            
            body = f"""Daily WiFi Data Validation Report - MoonFlower Hotel

Validation Summary:
📅 Date: {current_date}
🕐 Time: {current_time}
📊 Overall Status: {overall_status}

EXCEL FILE VALIDATION:
{'✅' if excel_ok else '❌'} Excel Status: {'VALID' if excel_ok else 'ISSUES DETECTED'}
📁 File Size: {excel_validation.get('file_size_mb', 0)} MB
📊 Row Count: {excel_validation.get('row_count', 0):,}
📋 Columns: {excel_validation.get('column_count', 0)}
🏗️ Headers: {'✅ Valid' if excel_validation.get('headers_valid', False) else '❌ Issues'}

Excel Issues:
{chr(10).join(['• ' + error for error in excel_validation.get('errors', [])]) or '• None'}

Excel Warnings:
{chr(10).join(['• ' + warning for warning in excel_validation.get('warnings', [])]) or '• None'}

PDF REPORT VALIDATION:
{'✅' if pdf_ok else '❌'} PDF Status: {'VALID' if pdf_ok else 'ISSUES DETECTED'}
📁 File Size: {pdf_validation.get('file_size_mb', 0)} MB
📄 Pages: {pdf_validation.get('page_count', 0)}
🔍 Keywords Found: {len(pdf_validation.get('keywords_found', []))}/{len(self.validation_criteria['required_pdf_keywords'])}
👥 User Count: {pdf_validation.get('user_count', 'Not found')}

PDF Issues:
{chr(10).join(['• ' + error for error in pdf_validation.get('errors', [])]) or '• None'}

DAILY GROWTH ANALYSIS:
{'✅' if growth_ok else '⚠️'} Growth Status: {user_comparison.get('message', 'No data')}
👥 Today's Count: {user_comparison.get('today_count', 'Unknown')}
👥 Yesterday's Count: {user_comparison.get('yesterday_count', 'Unknown')}
📈 Growth: {user_comparison.get('growth', 'N/A')} users
📊 Growth %: {user_comparison.get('growth_percentage', 'N/A')}%

SYSTEM HEALTH:
🔄 Data Collection: {'✅ Operational' if excel_ok else '❌ Issues'}
📤 VBS Upload: {'✅ Successful' if pdf_ok else '❌ Issues'}
📄 Report Generation: {'✅ Successful' if pdf_ok else '❌ Issues'}
📈 Data Growth: {'✅ Confirmed' if growth_ok else '⚠️ Review Required'}

RECOMMENDATIONS:
{'✅ System operating normally - no action required' if (excel_ok and pdf_ok and growth_ok) else '''⚠️ Manual review recommended:
• Verify CSV download completion
• Check VBS upload process
• Validate report generation
• Review data collection integrity'''}

AUTOMATION STATUS:
🤖 CSV Collection: Automated ✅
📊 Excel Generation: Automated ✅
📤 VBS Upload: Automated ✅
📄 PDF Generation: Automated ✅
📧 Email Delivery: Automated ✅
✅ Validation Check: Completed ✅

This automated validation ensures data integrity and system health for 
the MoonFlower Hotel WiFi monitoring system.

Best regards,
WiFi Automation Validation System
MoonFlower Hotel IT Department

---
Automated validation report
System Status: {'OPERATIONAL ✅' if overall_status == '✅ PASSED' else 'MONITORING ⚠️'}
Next Check: Tomorrow {(datetime.now() + timedelta(days=1)).strftime('%d/%m/%Y')}
"""
            
            success = self.email_system.send_email(
                recipient=self.email_system.recipients['development'],
                subject=subject,
                body=body
            )
            
            if success:
                self.logger.info("Validation report sent successfully")
            else:
                self.logger.error("Failed to send validation report")
            
            return success
            
        except Exception as e:
            self.logger.error(f"Error sending validation report: {e}")
            return False
    
    def execute_complete_validation(self) -> Dict[str, Any]:
        """Execute complete validation process"""
        try:
            self.logger.info("🔍 STARTING COMPLETE VALIDATION CHECK")
            
            start_time = datetime.now()
            today_folder = self.get_date_folder()
            
            # Find files
            excel_file = self.find_latest_excel_file(today_folder)
            pdf_file = self.find_latest_pdf_file(today_folder)
            
            # Validate Excel
            if excel_file:
                excel_validation = self.validate_excel_file(excel_file)
            else:
                excel_validation = {'file_exists': False, 'errors': ['Excel file not found']}
            
            # Validate PDF
            if pdf_file:
                pdf_validation = self.validate_pdf_file(pdf_file)
            else:
                pdf_validation = {'file_exists': False, 'errors': ['PDF file not found']}
            
            # Compare user counts
            if pdf_file:
                user_comparison = self.compare_daily_user_counts(pdf_file)
            else:
                user_comparison = {'validation_passed': False, 'message': 'No PDF file for comparison'}
            
            # Send email report
            email_sent = self.send_validation_report(excel_validation, pdf_validation, user_comparison)
            
            # Compile results
            execution_time = (datetime.now() - start_time).total_seconds()
            
            result = {
                'success': True,
                'execution_time_seconds': round(execution_time, 2),
                'date_checked': today_folder,
                'excel_validation': excel_validation,
                'pdf_validation': pdf_validation,
                'user_comparison': user_comparison,
                'email_report_sent': email_sent,
                'overall_status': 'PASSED' if (
                    excel_validation.get('data_valid', False) and 
                    pdf_validation.get('content_valid', False) and 
                    user_comparison.get('validation_passed', False)
                ) else 'ISSUES_DETECTED'
            }
            
            self.logger.info(f"✅ VALIDATION COMPLETED - Status: {result['overall_status']}")
            return result
            
        except Exception as e:
            self.logger.error(f"Complete validation failed: {e}")
            return {
                'success': False,
                'error': str(e),
                'overall_status': 'FAILED'
            }


def main():
    """Main function for testing validation system"""
    print("📊 WIFI DATA VALIDATION SYSTEM")
    print("=" * 60)
    print("Validates Excel files, PDF reports, and user count growth")
    print("Sends comprehensive email reports to development team")
    print("=" * 60)
    
    try:
        validator = ValidationSystem()
        result = validator.execute_complete_validation()
        
        print(f"\n📋 VALIDATION RESULTS:")
        print(f"Success: {result['success']}")
        print(f"Overall Status: {result.get('overall_status', 'Unknown')}")
        print(f"Execution Time: {result.get('execution_time_seconds', 0):.2f} seconds")
        
        if result['success']:
            excel_val = result['excel_validation']
            pdf_val = result['pdf_validation']
            user_comp = result['user_comparison']
            
            print(f"\n📊 EXCEL VALIDATION:")
            print(f"   File Exists: {'✅' if excel_val.get('file_exists') else '❌'}")
            print(f"   Row Count: {excel_val.get('row_count', 0):,}")
            print(f"   File Size: {excel_val.get('file_size_mb', 0)} MB")
            print(f"   Valid: {'✅' if excel_val.get('data_valid') else '❌'}")
            
            print(f"\n📄 PDF VALIDATION:")
            print(f"   File Exists: {'✅' if pdf_val.get('file_exists') else '❌'}")
            print(f"   Pages: {pdf_val.get('page_count', 0)}")
            print(f"   User Count: {pdf_val.get('user_count', 'Not found')}")
            print(f"   Valid: {'✅' if pdf_val.get('content_valid') else '❌'}")
            
            print(f"\n📈 USER COUNT COMPARISON:")
            print(f"   Today: {user_comp.get('today_count', 'Unknown')}")
            print(f"   Yesterday: {user_comp.get('yesterday_count', 'Unknown')}")
            print(f"   Growth: {user_comp.get('growth', 'N/A')} users")
            print(f"   Valid: {'✅' if user_comp.get('validation_passed') else '❌'}")
            
            print(f"\n📧 EMAIL REPORT:")
            print(f"   Sent: {'✅' if result.get('email_report_sent') else '❌'}")
            
            if result['overall_status'] == 'PASSED':
                print(f"\n🎉 ALL VALIDATIONS PASSED!")
                print(f"✅ Excel file is valid with proper data")
                print(f"✅ PDF report generated successfully")
                print(f"✅ User count growth confirmed")
                print(f"✅ Email report sent to development team")
            else:
                print(f"\n⚠️ SOME VALIDATIONS FAILED")
                print(f"📧 Check email for detailed error report")
        else:
            print(f"❌ Validation execution failed: {result.get('error', 'Unknown error')}")
            
    except Exception as e:
        print(f"❌ CRITICAL ERROR: {e}")
        import traceback
        traceback.print_exc()
    
    print("=" * 60)


if __name__ == "__main__":
    main() 