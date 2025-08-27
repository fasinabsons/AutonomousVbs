#!/usr/bin/env python3
"""
Outlook Automation for MoonFlower WiFi System
ONLY sends email to General Manager when PDF file exists
Clean, simple GM email delivery
"""

import win32com.client
import os
import time
import logging
from datetime import datetime, timedelta
from pathlib import Path
import sys

# Add parent directory for imports
sys.path.append(str(Path(__file__).parent.parent))

class OutlookAutomation:
    """Clean Outlook automation - ONLY for GM email when PDF exists"""
    
    def __init__(self):
        self.outlook = None
        self.logger = self._setup_logging()
        
        # CLEAN configuration - only what's needed
        self.config = {
            # FIXED: General Manager email  
            "gm_recipient": "loyed.basil@absons.ae",
            
            # FIXED: Sender account
            "sender_account": "mohamed.fasin@absons.ae",
            
            # Professional signature (blue bold formatting)
            "signature": """Best Regards<br>
<strong><span style="color: #0066CC; font-weight: bold;">Mohamed Fasin A F</span></strong><br>
<span style="color: black;">Software Developer</span><br><br>
<strong>E:</strong> <span style="color: #0066CC; font-weight: bold;">mohamed.fasin@absons.ae</span><br>
<strong>P:</strong> <span style="color: #0066CC; font-weight: bold;">+971 50 742 1288</span>"""
        }
        
        self.logger.info("Clean Outlook Automation initialized - GM email only")
    
    def _setup_logging(self):
        """Setup logging"""
        logger = logging.getLogger("OutlookGM")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # File handler
            try:
                date_folder = datetime.now().strftime("%d%b").lower()
                log_dir = Path(f"EHC_Logs/{date_folder}")
                log_dir.mkdir(parents=True, exist_ok=True)
                
                log_file = log_dir / f"gm_email_{datetime.now().strftime('%Y%m%d')}.log"
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
                
            except Exception as e:
                logger.warning(f"Could not setup file logging: {e}")
        
        return logger
    
    def connect_outlook(self):
        """Connect to Outlook"""
        try:
            self.logger.info("Connecting to Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.logger.info("✅ Connected to Outlook")
            return True
        except Exception as e:
            self.logger.error(f"❌ Outlook connection failed: {e}")
            return False
    
    def find_pdf_file(self):
        """Find YESTERDAY's PDF file ONLY - no email if yesterday PDF doesn't exist"""
        try:
            # ONLY look for yesterday's PDF (not today's)
            yesterday = datetime.now() - timedelta(days=1)
            date_folder = yesterday.strftime("%d%b").lower()
            pdf_dir = Path(f"EHC_Data_Pdf/{date_folder}")
            
            self.logger.info(f"Looking for YESTERDAY's PDF in: {date_folder}")
            
            if not pdf_dir.exists():
                self.logger.error(f"❌ Yesterday's PDF directory not found: {pdf_dir}")
                self.logger.info("💡 Solution: Run VBS Phase 1 → Phase 4 → Generate yesterday's PDF → Then send email")
                return None
            
            # Find PDF files
            pdf_files = list(pdf_dir.glob("*.pdf"))
            if not pdf_files:
                self.logger.error(f"❌ No PDF files in yesterday's folder: {pdf_dir}")
                self.logger.info("💡 Solution: Run VBS Phase 1 → Phase 4 → Generate yesterday's PDF → Then send email")
                return None
            
            # Return the latest PDF from yesterday
            latest_pdf = max(pdf_files, key=lambda x: x.stat().st_mtime)
            self.logger.info(f"✅ Found YESTERDAY's PDF: {latest_pdf.name}")
            return str(latest_pdf)
            
        except Exception as e:
            self.logger.error(f"❌ PDF search failed: {e}")
            return None
    
    def create_gm_email(self, pdf_file):
        """Create CLEAN GM email with PDF"""
        if not self.outlook:
            if not self.connect_outlook():
                return None
        
        try:
            # Create mail
            mail = self.outlook.CreateItem(0)
            
            # Set sender account (mohamed.fasin@absons.ae)
            accounts = self.outlook.Session.Accounts
            sender_account = None
            
            for account in accounts:
                if "mohamed.fasin@absons.ae" in account.SmtpAddress.lower():
                    sender_account = account
                    break
            
            if sender_account:
                mail.SendUsingAccount = sender_account
                self.logger.info(f"✅ Using sender: {sender_account.SmtpAddress}")
            else:
                self.logger.warning("⚠️ mohamed.fasin@absons.ae not found, using default")
            
            # Set recipient (GM)
            mail.To = self.config["gm_recipient"]
            
            # Generate current month/year
            current_date = datetime.now()
            month_names = [
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            ]
            month_year = f"{month_names[current_date.month - 1]} {current_date.year}"
            
            # Set subject
            mail.Subject = f"MoonFlower Active Users Count {month_year}"
            
            # Create CLEAN HTML body
            html_body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size: 14px; color: black; line-height: 1.6;">
                <p><strong>Dear Sir,</strong></p>
                
                <p>Good Morning, Please find the attached report detailing the active users for the Moon Flower project for {month_year}.</p>
                
                <p>Yours Kindly</p>
                
                <br>
                
                <p>
                    {self.config["signature"]}
                </p>
            </body>
            </html>
            """
            
            mail.HTMLBody = html_body
            
            # Add PDF attachment
            if pdf_file and os.path.exists(pdf_file):
                mail.Attachments.Add(pdf_file)
                self.logger.info(f"✅ PDF attached: {Path(pdf_file).name}")
            else:
                self.logger.error(f"❌ PDF not found: {pdf_file}")
                return None
            
            # NO high importance flag (user requested)
            
            self.logger.info(f"✅ GM email created")
            self.logger.info(f"To: {mail.To}")
            self.logger.info(f"Subject: {mail.Subject}")
            return mail
            
        except Exception as e:
            self.logger.error(f"❌ Email creation failed: {e}")
            return None
    
    def send_email(self, mail_object):
        """Send email immediately"""
        if not mail_object:
            self.logger.error("❌ No email to send")
            return False
        
        try:
            mail_object.Send()
            send_time = datetime.now().strftime("%H:%M:%S")
            self.logger.info(f"✅ GM email sent at {send_time}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Email send failed: {e}")
            return False
    
    def send_gm_email_if_pdf_exists(self):
        """MAIN FUNCTION: Only send GM email if PDF exists AND it's a weekday"""
        try:
            # Step 0: Check if today is a weekday (Monday-Friday only)
            today = datetime.now()
            day_of_week = today.weekday()  # 0=Monday, 6=Sunday
            
            if day_of_week >= 5:  # Saturday=5, Sunday=6
                day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                self.logger.info(f"⏰ Today is {day_names[day_of_week]} - GM emails only sent on weekdays")
                return False
            
            self.logger.info(f"✅ Weekday confirmed - proceeding with email")
            self.logger.info("🔍 Checking for PDF file...")
            
            # Step 1: Check for PDF
            pdf_file = self.find_pdf_file()
            if not pdf_file:
                self.logger.error("❌ NO PDF FOUND - Email not sent")
                return False
            
            # Step 2: Connect to Outlook
            if not self.connect_outlook():
                self.logger.error("❌ Outlook connection failed - Email not sent")
                return False
            
            # Step 3: Create email
            mail = self.create_gm_email(pdf_file)
            if not mail:
                self.logger.error("❌ Email creation failed - Email not sent")
                return False
            
            # Step 4: Send email
            success = self.send_email(mail)
            
            if success:
                self.logger.info("🎉 GM EMAIL SENT SUCCESSFULLY")
                self.logger.info(f"📧 To: {self.config['gm_recipient']}")
                self.logger.info(f"📎 Attachment: {Path(pdf_file).name}")
                return True
            else:
                self.logger.error("❌ Email sending failed")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ GM email process failed: {e}")
            return False

def main():
    """CLEAN main function - only send GM email if PDF exists"""
    print("📧 MOONFLOWER GM EMAIL AUTOMATION")
    print("=" * 50)
    print("✅ CLEAN VERSION: Only sends if PDF exists")
    print("✅ To: ramon.logan@absons.ae")
    print("✅ From: mohamed.fasin@absons.ae")
    print("=" * 50)
    
    outlook = OutlookAutomation()
    
    print("\n🔍 Checking for PDF file...")
    result = outlook.send_gm_email_if_pdf_exists()
    
    if result:
        print("\n🎉 SUCCESS: GM email sent!")
        print("✅ PDF attached")
        print("✅ Professional formatting")
        print("✅ Correct sender/recipient")
    else:
        print("\n❌ FAILED: Email not sent")
        print("⚠️ Check if PDF file exists")
        print("⚠️ Check Outlook connection")
        print("⚠️ Check sender account setup")
    
    return result

if __name__ == "__main__":
    main() 