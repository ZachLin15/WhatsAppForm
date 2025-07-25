import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from tqdm import tqdm
import time
import logging

logging.basicConfig(filename=r'C:\Users\USER\ImportOracle\pythonProject1\SendFonterraEmail.log',
                    level=logging.INFO,  # Set log level to DEBUG
                    format='%(asctime)s - %(levelname)s - %(message)s')

def send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body, attachment_path):
    """Sends an email with an attachment."""

    # Create message object
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject

    # Attach body to email
    message.attach(MIMEText(body, 'plain'))  # Use 'html' for HTML content

    # Attach file
    if attachment_path:  # Only attach if a path is provided
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)

                # Add header with filename
                filename = os.path.basename(attachment_path)
                part.add_header("Content-Disposition", f"attachment; filename= {filename}")

                message.attach(part)
        except FileNotFoundError:
            logging.error(f"Error: Attachment file not found: {attachment_path}")
            return False  # Indicate failure

    try:
        # Connect to SMTP server
        with smtplib.SMTP("smtp.office365.com", 587) as server: # Or smtp.office365.com
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, receiver_email, message.as_string())
        logging.info(f"Email sent successfully to {receiver_email}!")
        return True # Indicate success
    except Exception as e:
        logging.error(f"Error sending email {receiver_email}: {e}")
        return False # Indicate failure

# Example usage:
sender_email = "admin1@lshworld.com" # Replace with your email
sender_password = "dpvqmxwsrxvxmbvr"  # Replace with your password or app password
receiver_email = ['leezhenglin95@gmail.com','amore@lshworld.com','EeChing.Tan@fonterra.com','ADIBAHAIDA.YAHAYA@fonterra.com','Gail.Chong@fonterra.com','Sheldon.Yen@fonterra.com','meiyu.chen2@fonterra.com','JiaJia.Poh@fonterra.com','Darren.Ho@fonterra.com']
#receiver_email = ['leezhenglin95@gmail.com','zhenglin@limsianghuat.com']
subject = "LimSiangHuat x Fonterra Daily Report"
body = """Dear Sir/Madam,

Please find attached the latest Sales Report.

This report has been automatically generated. For any inquiries, please contact us via email at amore@lshworld.com.

Best Regards,
Amorelle Ong
Lim Siang Huat Pte Ltd
6 Fishery Port Road, #02M, Singapore 619747
T: 62647592 | F: 62624144 | HP: 93662630 | W: www.LimSiangHuat.com | e-store: www.LimSiangHuat.com/shop"""

#attachment_path = "path/to/your/file.pdf"  # Replace with the actual path
#If you don't want to add attachement, set attachment_path to None, or ""

#Run the batch file to run the query
os.startfile(r"C:\Users\USER\Documents\UiPath\Published Prject\Run.Fonterra.VB.1.0.3.bat")

#Give refresh process 20 Sec
for _ in tqdm(range(20), desc="Refreshing", unit="s"):
    time.sleep(1)

for files in os.listdir(r"W:\FRSALES\Output"):



    files = os.path.join("W:\\FRSALES\\Output", files)
    if files.__contains__("DS"):
        for emails in receiver_email:
            send_email_with_attachment(sender_email, sender_password, emails, subject, body, files)
        os.remove(files)
        logging.info(f"{files} Removed")
