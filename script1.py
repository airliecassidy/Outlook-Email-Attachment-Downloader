"""
Automated Email Attachment Downloader for Outlook
Monitors Romspen Investor Relations folder for weekly emails and downloads attachments

Need to change:
sender_email
shared_folder_name
download_folder
to the actual values

"""
import os
import sys
import time
import logging
import smtplib
from datetime import datetime, timedelta
import json
import hashlib
import win32com.client
from typing import Optional, List, Tuple
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class OutlookAttachmentDownloader:

    def __init__(self, config_file: str = "config.json"):
        self.config = self.load_config(config_file)
        self._setup_logging()
        self.outlook = None
        self.namespace = None
        self.processed_emails = self.load_processed_emails()

    def load_config(self, config_file: str) -> dict:
        default_config = {
            "sender_email": "sender@email.com",  # email address of weekly sender
            "shared_folder_name": "Romspen Investor Relations",  # inbox to be scanned
            "download_folder": r"\\shared\drive\path\attachments",  # pathway to shared download folder
            "expected_day": 1,  # Monday = 0, Tuesday = 1, etc.
            "notification_email": "admin@company.com",
            "smtp_server": "smtp.company.com",
            "smtp_port": 587,
            "smtp_username": "notifications@company.com",
            "smtp_password": "password",
            "check_interval_minutes": 30,
            "processed_emails_file": "processed_emails.json"
        }

        try:
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    user_config = json.load(f)
                    default_config.update(user_config)
            else:
                with open(config_file, 'w') as f:
                    json.dump(default_config, f, indent=4)
                self._log_info(f"Created default config file: {config_file}")
        except Exception as e:
            self._log_info(f"Error loading config: {e}")

        return default_config

    def _setup_logging(self):
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler('outlook_downloader.log'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)

    def _log_info(self, message: str):
        self.logger.info(message)

    def _log_error(self, message: str):
        """Log error message"""
        self.logger.error(message)

    def load_processed_emails(self) -> set:

        try:
            if os.path.exists(self.config["processed_email_file"]):
                with open(self.config["processed_emails_file"], 'r') as f:
                    return set(json.load(f))
        except Exception as e:
            self._log_error(f"Error loading processed emails: {e}")
        return set()

    def save_processed_emails(self):
        try:
            with open(self.config["processed_emails_file"], 'w') as f:
                json.dump(list(self.processed_emails), f)
        except Exception as e:
            self._log_error(f"Error saving processed emails: {e}")

    def _generate_email_id(self, email_item) -> str:
        """Generate unique ID for email based on sender, subject, and received time"""
        try:
            content = f"{email_item.SenderEmailAddress}{email_item.Subject}{email_item.ReceivedTime}"
            return hashlib.md5(content.encode()).hexdigest()
        except:
            return str(hash(str(email_item)))

    def connect_to_outlook(self) -> bool:
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self._log_info("Successfully connected to Outlook")
            return True
        except Exception as e:
            self._log_error(f"Failed to connect to Outlook: {e}")
            return False

    def find_shared_folder(self) -> Optional[object]:  # finding the romspen shared folder
        try:
            for store in self.namespace.Stores:
                if "Romspen" in store.DisplayName or "Investor Relations" in store.DisplayName:
                    root_folder = store.GetRootFolder()
                    for folder in root_folder.Folders:
                        if folder.Name == self.config["shared_folder_name"]:
                            self._log_info(f"Found shared folder: {folder.Name}")
                            return folder

            inbox = self.namespace.GetDefaultFolder(6)  # Inbox
            for folder in inbox.Folders:
                if folder.Name == self.config["shared_folder_name"]:
                    self._log_info(f"Found folder in inbox: {folder.Name}")
                    return folder

            self._log_error(f"Shared folder '{self.config['shared_folder_name']}' not found")
            return None

        except Exception as e:
            self._log_error(f"Error finding shared folder: {e}")
            return None

    def get_weekly_emails(self, folder, target_date: datetime) -> List[object]:

        emails = []

        try:
            start_date = target_date.replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=1)

            for item in folder.Items:
                try:  # try to make this cleaner bru
                    if (hasattr(item, 'SenderEmailAddress') and hasattr(item,
                                                                        'ReceivedTime') and item.SenderEmailAddress.lower() ==
                            self.config["sender_email"].lower()):
                        received_time = item.ReceivedTime

                        if isinstance(received_time, str):
                            continue
                        if start_date <= received_time < end_date:
                            email_id = self._generate_email_id(item)
                        if email_id not in self.processed_emails:
                            emails.append(item)
                            self._log_info(f"Found new email from {item.SenderEmailAddress} at {received_time}")
                except Exception as e:
                    self._log_error(f"Error processing email item: {e}")
                    continue
        except Exception as e:
            self._log_error(f"Error retrieving emails: {e}")

        return emails

    def download_attachments(self, email_item) -> bool:
        try:
            if not os.path.exists(self.config["download_folder"]):
                os.makedirs(self.config["download_folder"])
                self._log_info(f"Created download folder: {self.config['download_folder']}")

            attachments_downloaded = 0

            for attachment in email_item.Attachments:
                try:
                    filename = attachment.FileName
                    filepath = os.path.join(self.config["download_folder"], filename)

                    # Avoid overwriting existing files by adding timestamp
                    if os.path.exists(filepath):
                        name, ext = os.path.splitext(filename)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"{name}_{timestamp}{ext}"
                        filepath = os.path.join(self.config["download_folder"], filename)

                    attachment.SaveAsFile(filepath)
                    attachments_downloaded += 1
                    self._log_info(f"Downloaded attachment: {filename}")

                except Exception as e:
                    self._log_error(f"Error downloading attachment {attachment.FileName}: {e}")

            if attachments_downloaded > 0:
                email_id = self._generate_email_id(email_item)
                self.processed_emails.add(email_id)
                self._save_processed_emails()
                return True
            else:
                self._log_info("No attachments found in email")
                return False

        except Exception as e:
            self._log_error(f"Error in download_attachments: {e}")
            return False

    def send_notification(self, message: str, subject: str = "Email Attachment Downloader Alert"):
        """Send email notification"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.config["smtp_username"]
            msg['To'] = self.config["notification_email"]
            msg['Subject'] = subject

            msg.attach(MIMEText(message, 'plain'))

            server = smtplib.SMTP(self.config["smtp_server"], self.config["smtp_port"])
            server.starttls()
            server.login(self.config["smtp_username"], self.config["smtp_password"])
            text = msg.as_string()
            server.sendmail(self.config["smtp_username"], self.config["notification_email"], text)
            server.quit()

            self._log_info(f"Notification sent: {subject}")

        except Exception as e:
            self._log_error(f"Failed to send notification: {e}")

    def check_for_weekly_email(self):
        """Main function to check for weekly email and download attachments"""
        try:
            if not self.connect_to_outlook():
                return

            shared_folder = self.find_shared_folder()
            if not shared_folder:
                return

            today = datetime.now()
            days_since_target = (today.weekday() - self.config["expected_day"]) % 7
            target_date = today - timedelta(days=days_since_target)

            if today.weekday() >= self.config["expected_day"]:
                emails = self.get_weekly_emails(shared_folder, target_date)

                if emails:
                    for email in emails:
                        success = self.download_attachments(email)
                        if success:
                            self._log_info(f"Successfully processed email from {email.SenderEmailAddress}")
                        else:
                            self._log_error(f"Failed to download attachments from {email.SenderEmailAddress}")
                else:
                    if (today - target_date).days >= 1:
                        notification_msg = (f"Weekly email from {self.config['sender_email']} "
                                            f"was not received on {target_date.strftime('%Y-%m-%d')} "
                                            f"(Expected on {target_date.strftime('%A')})")
                        self.send_notification(notification_msg, "Missing Weekly Email Alert")

        except Exception as e:
            self._log_error(f"Error in check_for_weekly_email: {e}")
            error_msg = f"System error occurred: {str(e)}"
            self.send_notification(error_msg, "Email Downloader System Error")

    def run_continuous(self):
        """Run the downloader continuously"""
        self._log_info("Starting continuous email attachment downloader...")

        while True:
            try:
                self.check_for_weekly_email()
                sleep_seconds = self.config["check_interval_minutes"] * 60
                self._log_info(f"Check completed. Sleeping for {self.config['check_interval_minutes']} minutes...")
                time.sleep(sleep_seconds)

            except KeyboardInterrupt:
                self._log_info("Received shutdown signal. Stopping downloader...")
                break
            except Exception as e:
                self._log_error(f"Unexpected error in main loop: {e}")
                time.sleep(300)  # Sleep 5 minutes on error before retrying


def main():
    """Main entry point"""
    try:
        downloader = OutlookAttachmentDownloader()
        downloader.run_continuous()
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()