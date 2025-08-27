#!/usr/bin/env python3
"""
Simple Email Delivery System for MoonFlower WiFi Automation
ONLY sends notifications to you (faseenm@gmail.com)
Clean separation: GM emails handled by outlook_automation.py
"""

import smtplib
import sys
import logging
from datetime import datetime
from pathlib import Path
from email.mime.text import MIMEText

class SimpleEmailDelivery:
    """Simple notification system - ONLY to faseenm@gmail.com"""
    
    def __init__(self):
        self.logger = self._setup_logging()
        
        # FIXED Gmail configuration for notifications
        self.gmail_config = {
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'sender_email': 'fasin.absons@gmail.com',
            'sender_password': 'zrxj vfjt wjos wkwy',
            'recipient': 'faseenm@gmail.com'
        }
        
        self.logger.info("Simple Notification System initialized")
    
    def _setup_logging(self):
        """Setup simple logging"""
        logger = logging.getLogger("EmailNotifications")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            console_handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
        
        return logger
    
    def send_notification(self, subject, message):
        """Send simple notification email to you only"""
        try:
            current_time = datetime.now().strftime("%H:%M")
            
            body = f"""Hi Fasin,

{message}

Time: {current_time}
Date: {datetime.now().strftime('%d/%m/%Y')}

MoonFlower Automation"""
            
            # Create email
            msg = MIMEText(body)
            msg['Subject'] = subject
            msg['From'] = self.gmail_config['sender_email']
            msg['To'] = self.gmail_config['recipient']
            
            # Send via Gmail SMTP
            with smtplib.SMTP(self.gmail_config['smtp_server'], self.gmail_config['smtp_port']) as server:
                server.starttls()
                server.login(self.gmail_config['sender_email'], self.gmail_config['sender_password'])
                server.send_message(msg)
            
            self.logger.info(f"✅ Notification sent: {subject}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Notification failed: {e}")
            return False
    
    def notify_csv_complete(self, files_count, date_folder):
        """Notify when CSV files are complete"""
        subject = f"CSV Downloads Complete - {files_count} Files"
        message = f"""CSV downloads completed successfully.

Files Downloaded: {files_count}
Location: EHC_Data/{date_folder}/

Networks covered:
- EHC TV Network
- EHC-15 Network  
- Reception Hall-Mobile
- Reception Hall-TV Network

Next: Excel merge will start automatically."""
        
        return self.send_notification(subject, message)
    
    def notify_excel_complete(self, excel_file, records_count):
        """Notify when Excel merge is complete"""
        excel_path = Path(excel_file)
        subject = f"Excel Merge Complete - {records_count:,} Records"
        message = f"""Excel merge completed successfully.

File: {excel_path.name}
Records: {records_count:,}
Location: {excel_path.parent}

Data processed and ready for VBS upload.
Next: VBS processing will start."""
        
        return self.send_notification(subject, message)
    
    def notify_upload_complete(self, upload_duration_hours):
        """Notify when VBS upload is complete"""
        subject = "VBS Upload Complete"
        message = f"""VBS upload process completed successfully.

Duration: {upload_duration_hours:.1f} hours
Process: Excel data uploaded to VBS system

Upload successful - completion popup detected.
VBS software closed automatically.
Next: PDF report generation will start."""
        
        return self.send_notification(subject, message)
    
    def notify_pdf_created(self, pdf_file):
        """Notify when PDF report is created"""
        pdf_path = Path(pdf_file)
        subject = "PDF Report Created"
        message = f"""PDF report generated successfully.

File: {pdf_path.name}
Location: {pdf_path.parent}

Report ready for next step."""
        
        return self.send_notification(subject, message)
    
    def notify_automation_complete(self):
        """Notify when complete automation cycle is finished"""
        subject = "Automation Cycle Complete"
        message = """Complete automation cycle finished successfully.

✅ CSV downloads completed
✅ Excel merge completed  
✅ VBS upload completed
✅ PDF report created

Daily workflow completed successfully.
All tasks finished without issues."""
        
        return self.send_notification(subject, message)
    
    def notify_csv_only_complete(self, files_count, date_folder):
        """Notify when CSV files are complete but Excel not merged yet"""
        subject = f"CSV Downloads Complete - {files_count} Files (Excel Pending)"
        message = f"""CSV downloads completed successfully.

Files Downloaded: {files_count}
Location: EHC_Data/{date_folder}/
Status: Excel merge pending (will run at 12:35 PM)

Networks covered:
- EHC TV Network
- EHC-15 Network  
- Reception Hall-Mobile
- Reception Hall-TV Network

Next: Excel merge scheduled for 12:35 PM."""
        
        return self.send_notification(subject, message)
    
    def notify_csv_failed(self):
        """Notify when CSV downloads failed"""
        subject = "CSV Downloads Failed - Action Required"
        message = """CSV downloads failed after 10 retry attempts.

❌ All download attempts unsuccessful
❌ Unable to connect to Ruckus Controller
❌ Network or authentication issues possible

Manual intervention required.
Check network connectivity and controller status."""
        
        return self.send_notification(subject, message)


def send_notification(notification_type, *args):
    """Quick function to send notifications from other scripts"""
    email_system = SimpleEmailDelivery()
    
    if notification_type == "csv_complete":
        return email_system.notify_csv_complete(args[0], args[1])
    elif notification_type == "csv_only_complete":
        return email_system.notify_csv_only_complete(args[0], args[1])
    elif notification_type == "csv_failed":
        return email_system.notify_csv_failed()
    elif notification_type == "excel_complete":
        return email_system.notify_excel_complete(args[0], args[1])
    elif notification_type == "upload_complete":
        return email_system.notify_upload_complete(args[0])
    elif notification_type == "pdf_created":
        return email_system.notify_pdf_created(args[0])
    elif notification_type == "automation_complete":
        return email_system.notify_automation_complete()
    else:
        return False


def main():
    """Test the notification system"""
    if len(sys.argv) > 1:
        notification_type = sys.argv[1]
        args = sys.argv[2:] if len(sys.argv) > 2 else []
        result = send_notification(notification_type, *args)
        print(f"Notification sent: {result}")
    else:
        print("Usage: python email_delivery.py <notification_type> [args]")
        print("Types: csv_complete, csv_only_complete, csv_failed, excel_complete, upload_complete, pdf_created, automation_complete")


if __name__ == "__main__":
    main()