#!/usr/bin/env python3
"""
Enhanced Email Reporting System
Handles automated PDF report delivery with intelligent scheduling
"""

import smtplib
import os
import json
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Any
import calendar

class EnhancedEmailReporting:
    def __init__(self, config_path: str = "config/email_settings.json"):
        self.logger = self._setup_logging()
        self.config = self._load_config(config_path)
        self.current_date = datetime.now()
        
    def _setup_logging(self) -> logging.Logger:
        logger = logging.getLogger("EmailReporting")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(console_formatter)
            logger.addHandler(console_handler)
            
            try:
                log_file = "EHC_Logs/email_reporting.log"
                os.makedirs(os.path.dirname(log_file), exist_ok=True)
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                file_handler.setFormatter(file_formatter)
                logger.addHandler(file_handler)
            except Exception:
                pass
        
        return logger
    
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        """Load email configuration"""
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r') as f:
                    return json.load(f)
            else:
                return self._create_default_config(config_path)
        except Exception as e:
            self.logger.error(f"Config loading failed: {e}")
            return self._get_fallback_config()
    
    def _create_default_config(self, config_path: str) -> Dict[str, Any]:
        """Create default email configuration"""
        default_config = {
            "smtp_settings": {
                "server": "smtp.office365.com",
                "port": 587,
                "use_tls": True,
                "username": "automation@company.com",
                "password": "app_password_here"
            },
            "recipients": {
                "general_manager": "gm@company.com",
                "operations": ["ops1@company.com", "ops2@company.com"],
                "management": ["manager1@company.com", "manager2@company.com"]
            },
            "scheduling": {
                "send_weekdays": True,
                "send_weekends": False,
                "friday_hold_until_monday": True,
                "monday_sends_sunday_report": True
            },
            "sender_info": {
                "name": "MoonFlower Automation",
                "email": "automation@company.com"
            }
        }
        
        # Save default config
        try:
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w') as f:
                json.dump(default_config, f, indent=2)
        except Exception as e:
            self.logger.error(f"Failed to save default config: {e}")
        
        return default_config
    
    def _get_fallback_config(self) -> Dict[str, Any]:
        """Fallback configuration"""
        return {
            "smtp_settings": {
                "server": "smtp.gmail.com",
                "port": 587,
                "use_tls": True,
                "username": "fallback@gmail.com",
                "password": "fallback_password"
            },
            "recipients": {
                "general_manager": "admin@localhost",
                "operations": ["admin@localhost"],
                "management": ["admin@localhost"]
            },
            "scheduling": {
                "send_weekdays": True,
                "send_weekends": False,
                "friday_hold_until_monday": True,
                "monday_sends_sunday_report": True
            }
        }
    
    def should_send_report_today(self) -> Dict[str, Any]:
        """Determine if report should be sent today based on scheduling logic"""
        today = self.current_date
        weekday = today.weekday()  # 0=Monday, 6=Sunday
        
        result = {
            "should_send": False,
            "report_date": today,
            "reason": "",
            "recipient_type": "operations"
        }
        
        # Monday logic - send Sunday's report to GM
        if weekday == 0:  # Monday
            if self.config["scheduling"]["monday_sends_sunday_report"]:
                result.update({
                    "should_send": True,
                    "report_date": today - timedelta(days=1),  # Sunday
                    "reason": "Monday - sending Sunday's report",
                    "recipient_type": "general_manager"
                })
        
        # Tuesday-Thursday logic - send previous day's report
        elif weekday in [1, 2, 3]:  # Tuesday, Wednesday, Thursday
            result.update({
                "should_send": True,
                "report_date": today - timedelta(days=1),
                "reason": f"Weekday - sending previous day's report",
                "recipient_type": "general_manager"
            })
        
        # Friday logic - hold report unless configured otherwise
        elif weekday == 4:  # Friday
            if not self.config["scheduling"]["friday_hold_until_monday"]:
                result.update({
                    "should_send": True,
                    "report_date": today - timedelta(days=1),  # Thursday
                    "reason": "Friday - sending Thursday's report (no hold configured)",
                    "recipient_type": "general_manager"
                })
            else:
                result.update({
                    "should_send": False,
                    "reason": "Friday - holding report until Monday"
                })
        
        # Weekend logic - generate but don't send unless configured
        elif weekday in [5, 6]:  # Saturday, Sunday
            if self.config["scheduling"]["send_weekends"]:
                result.update({
                    "should_send": True,
                    "report_date": today - timedelta(days=1),
                    "reason": "Weekend sending enabled",
                    "recipient_type": "operations"
                })
            else:
                result.update({
                    "should_send": False,
                    "reason": "Weekend - reports generated but not sent"
                })
        
        self.logger.info(f"Send decision: {result}")
        return result
    
    def find_latest_pdf_report(self, target_date: datetime) -> Optional[str]:
        """Find the latest PDF report for the specified date"""
        try:
            date_folder = target_date.strftime("%d%b").lower()
            pdf_folder = f"EHC_Data_Pdf/{date_folder}"
            
            if not os.path.exists(pdf_folder):
                self.logger.warning(f"PDF folder not found: {pdf_folder}")
                return None
            
            # Look for PDF files matching the pattern
            pdf_files = list(Path(pdf_folder).glob("*.pdf"))
            
            if not pdf_files:
                self.logger.warning(f"No PDF files found in {pdf_folder}")
                return None
            
            # Get the most recent PDF file
            latest_pdf = max(pdf_files, key=os.path.getctime)
            
            self.logger.info(f"Found PDF report: {latest_pdf}")
            return str(latest_pdf)
            
        except Exception as e:
            self.logger.error(f"PDF search failed: {e}")
            return None
    
    def generate_monthly_signature(self) -> str:
        """Generate signature with monthly rotation"""
        month_signatures = {
            1: "Best regards for the New Year",
            2: "With warm regards this February", 
            3: "Spring greetings",
            4: "Best wishes this April",
            5: "Warm May regards",
            6: "Mid-year greetings",
            7: "Summer regards",
            8: "Late summer wishes",
            9: "Autumn greetings",
            10: "Fall regards",
            11: "Pre-holiday wishes",
            12: "Year-end regards"
        }
        
        current_month = self.current_date.month
        monthly_greeting = month_signatures.get(current_month, "Best regards")
        
        signature = f"""
{monthly_greeting},

{self.config['sender_info']['name']}
MoonFlower Automation System
Generated on {self.current_date.strftime('%d %B %Y')}
"""
        
        return signature
    
    def create_report_email(self, pdf_path: str, report_date: datetime, recipient_type: str) -> MIMEMultipart:
        """Create formatted email with PDF attachment"""
        msg = MIMEMultipart()
        
        # Email headers
        msg['From'] = self.config['sender_info']['email']
        msg['Subject'] = f"📊 MoonFlower Daily Report - {report_date.strftime('%d/%m/%Y')}"
        
        # Set recipients based on type
        if recipient_type == "general_manager":
            msg['To'] = self.config['recipients']['general_manager']
        elif recipient_type == "management":
            msg['To'] = ', '.join(self.config['recipients']['management'])
        else:
            msg['To'] = ', '.join(self.config['recipients']['operations'])
        
        # Email body - formal and concise
        body = f"""
Dear Team,

Please find attached the MoonFlower active users report for {report_date.strftime('%d %B %Y')}. All systems processed successfully with complete data coverage.

Report summary: Daily WiFi user activity analysis with consolidated access point data and usage metrics.

{self.generate_monthly_signature()}
        """
        
        msg.attach(MIMEText(body.strip(), 'plain'))
        
        # Attach PDF if exists
        if pdf_path and os.path.exists(pdf_path):
            try:
                with open(pdf_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                
                encoders.encode_base64(part)
                
                filename = os.path.basename(pdf_path)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {filename}'
                )
                
                msg.attach(part)
                self.logger.info(f"PDF attached: {filename}")
                
            except Exception as e:
                self.logger.error(f"PDF attachment failed: {e}")
        
        return msg
    
    def send_email(self, msg: MIMEMultipart) -> bool:
        """Send email using SMTP"""
        try:
            smtp_config = self.config['smtp_settings']
            
            # Create SMTP session
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
            
            if smtp_config.get('use_tls', True):
                server.starttls()
            
            # Login with credentials
            server.login(smtp_config['username'], smtp_config['password'])
            
            # Send email
            text = msg.as_string()
            server.sendmail(msg['From'], msg['To'].split(', '), text)
            server.quit()
            
            self.logger.info(f"Email sent successfully to: {msg['To']}")
            return True
            
        except Exception as e:
            self.logger.error(f"Email sending failed: {e}")
            return False
    
    def send_daily_report(self) -> Dict[str, Any]:
        """Main method to send daily report"""
        try:
            result = {
                "success": False,
                "timestamp": self.current_date.isoformat(),
                "message": "",
                "recipients": [],
                "pdf_attached": False
            }
            
            # Check if we should send today
            send_decision = self.should_send_report_today()
            
            if not send_decision["should_send"]:
                result.update({
                    "success": True,
                    "message": send_decision["reason"]
                })
                return result
            
            # Find PDF report for the target date
            report_date = send_decision["report_date"]
            pdf_path = self.find_latest_pdf_report(report_date)
            
            if not pdf_path:
                result.update({
                    "message": f"No PDF report found for {report_date.strftime('%d/%m/%Y')}"
                })
                return result
            
            # Create and send email
            recipient_type = send_decision["recipient_type"]
            email_msg = self.create_report_email(pdf_path, report_date, recipient_type)
            
            if self.send_email(email_msg):
                result.update({
                    "success": True,
                    "message": f"Report sent successfully for {report_date.strftime('%d/%m/%Y')}",
                    "recipients": email_msg['To'].split(', '),
                    "pdf_attached": True
                })
            else:
                result.update({
                    "message": "Email sending failed"
                })
            
            return result
            
        except Exception as e:
            self.logger.error(f"Report sending failed: {e}")
            return {
                "success": False,
                "message": f"Error: {str(e)}",
                "timestamp": self.current_date.isoformat()
            }
    
    def send_notification_email(self, subject: str, message: str, recipients: List[str], use_emojis: bool = True) -> bool:
        """Send casual notification email with emojis"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.config['sender_info']['email']
            msg['To'] = ', '.join(recipients)
            
            if use_emojis:
                msg['Subject'] = f"🤖 {subject}"
                emoji_message = f"👋 Hi team!\n\n{message}\n\n🔧 MoonFlower Automation"
            else:
                msg['Subject'] = subject
                emoji_message = message
            
            msg.attach(MIMEText(emoji_message, 'plain'))
            
            return self.send_email(msg)
            
        except Exception as e:
            self.logger.error(f"Notification email failed: {e}")
            return False

def main():
    """Test the email reporting system"""
    reporter = EnhancedEmailReporting()
    result = reporter.send_daily_report()
    print(f"Report result: {result}")

if __name__ == "__main__":
    main()