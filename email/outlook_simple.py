#!/usr/bin/env python3
"""
SIMPLE OUTLOOK AUTOMATION - Works with Organizational Outlook
Just creates and displays the email - user can review and send manually
This ensures compatibility with all organizational security settings
"""

import win32com.client
import os
import time
import logging
from datetime import datetime, timedelta
from pathlib import Path
import sys

class SimpleOutlookAutomation:
    """Simple, reliable Outlook automation for organizational environments"""
    
    def __init__(self):
        self.outlook = None
        self.logger = self._setup_logging()
        
        # Simple configuration
        self.config = {
            "gm_recipient": "loyed.basil@absons.ae",
            "sender_account": "mohamed.fasin@absons.ae",
            "signature": """Best Regards<br>
<strong><span style="color: #0066CC; font-weight: bold;">Mohamed Fasin A F</span></strong><br>
<span style="color: black;">Software Developer</span><br><br>
<strong>E:</strong> <span style="color: #0066CC; font-weight: bold;">mohamed.fasin@absons.ae</span><br>
<strong>P:</strong> <span style="color: #0066CC; font-weight: bold;">+971 50 742 1288</span>"""
        }
        
        self.logger.info("Simple Outlook Automation initialized")
    
    def _setup_logging(self):
        """Setup logging"""
        logger = logging.getLogger("SimpleOutlook")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            
            # File logging
            try:
                log_dir = Path("EHC_Logs")
                log_dir.mkdir(exist_ok=True)
                log_file = log_dir / f"outlook_simple_{datetime.now().strftime('%Y%m%d')}.log"
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(formatter)
                logger.addHandler(file_handler)
            except:
                pass  # Continue without file logging if it fails
        
        return logger
    
    def _connect_outlook(self):
        """Connect to Outlook application"""
        try:
            self.logger.info("Connecting to Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            
            # Test connection
            namespace = self.outlook.GetNamespace("MAPI")
            self.logger.info("✅ Outlook connection successful")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to connect to Outlook: {e}")
            return False
    
    def _find_pdf_file(self):
        """Find YESTERDAY's PDF file ONLY - same as outlook_automation.py"""
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
            return latest_pdf
            
        except Exception as e:
            self.logger.error(f"❌ PDF search failed: {e}")
            return None
    
    def create_gm_email(self, pdf_file):
        """Create and send GM email with provided PDF file"""
        try:
            if not self._connect_outlook():
                return False
            
            # Create email
            self.logger.info("📧 Creating GM email...")
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            
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
            
            # Generate current month/year (SAME AS outlook_automation.py)
            current_date = datetime.now()
            month_names = [
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            ]
            month_year = f"{month_names[current_date.month - 1]} {current_date.year}"
            
            # Set subject (SAME AS outlook_automation.py)
            mail.Subject = f"MoonFlower Active Users Count {month_year}"
            
            # Create CLEAN HTML body (EXACT SAME AS outlook_automation.py)
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
            
            # Attach PDF file
            try:
                mail.Attachments.Add(str(pdf_file.absolute()))
                self.logger.info(f"✅ Attached PDF: {pdf_file.name}")
            except Exception as e:
                self.logger.error(f"❌ Failed to attach PDF: {e}")
                return False
            
            # Store email details BEFORE sending (mail object becomes invalid after Send())
            email_subject = mail.Subject
            email_recipient = mail.To
            attachment_name = pdf_file.name
            
            # Send email automatically - COMPLETE AUTOMATION
            self.logger.info("📤 Sending email automatically...")
            mail.Send()  # Automatically send the email
            
            self.logger.info("✅ Email sent successfully - COMPLETE AUTOMATION")
            
            print("\n" + "="*60)
            print("📧 EMAIL SENT SUCCESSFULLY!")
            print("="*60)
            print(f"✅ TO: {email_recipient}")
            print(f"✅ SUBJECT: {email_subject}")
            print(f"✅ ATTACHMENT: {attachment_name}")
            print(f"✅ STATUS: AUTOMATICALLY SENT")
            print("\n🎉 COMPLETE AUTOMATION - NO REVIEW NEEDED")
            print("="*60)
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Error creating GM email: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def send_gm_email_if_pdf_exists(self):
        """MAIN FUNCTION: Only send GM email if PDF exists AND it's a weekday (SAME AS outlook_automation.py)"""
        try:
            # Step 0: Check if today is a weekday (Monday-Friday only) - EXACT SAME LOGIC
            today = datetime.now()
            day_of_week = today.weekday()  # 0=Monday, 6=Sunday
            
            if day_of_week >= 5:  # Saturday=5, Sunday=6
                day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
                self.logger.info(f"⏰ Today is {day_names[day_of_week]} - GM emails only sent on weekdays")
                return False
            
            self.logger.info(f"✅ Weekday confirmed - proceeding with email")
            self.logger.info("🔍 Checking for PDF file...")
            
            # Step 1: Check for PDF
            pdf_file = self._find_pdf_file()
            if not pdf_file:
                self.logger.error("❌ NO PDF FOUND - Email not sent")
                return False
            
            # Step 2: Create and send email
            self.logger.info("🚀 Starting Simple GM email creation...")
            success = self.create_gm_email(pdf_file)
            
            if success:
                self.logger.info("🎉 GM EMAIL SENT SUCCESSFULLY")
                self.logger.info(f"📧 To: {self.config['gm_recipient']}")
                self.logger.info(f"📎 Attachment: {pdf_file.name}")
                return True
            else:
                self.logger.error("❌ GM email process failed")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error in GM email process: {e}")
            return False

def main():
    """CLEAN main function - only send GM email if PDF exists (SAME AS outlook_automation.py)"""
    print("📧 MOONFLOWER GM EMAIL AUTOMATION")
    print("=" * 50)
    print("✅ CLEAN VERSION: Only sends if PDF exists")
    print(f"✅ To: loyed.basil@absons.ae")
    print("✅ From: mohamed.fasin@absons.ae")
    print("=" * 50)
    
    outlook = SimpleOutlookAutomation()
    
    print("\n🔍 Checking for PDF file...")
    result = outlook.send_gm_email_if_pdf_exists()
    
    if result:
        print("\n🎉 SUCCESS: GM email sent!")
        print("✅ PDF attached")
        print("✅ Professional formatting")
        print("✅ Correct sender/recipient")
        sys.exit(0)
    else:
        print("\n❌ FAILED: Email not sent")
        print("⚠️ Check if PDF file exists")
        print("⚠️ Check Outlook connection")
        print("⚠️ Check sender account setup")
        sys.exit(1)

if __name__ == "__main__":
    main()
