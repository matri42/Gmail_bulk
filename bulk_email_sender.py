import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
from dotenv import load_dotenv
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(
    filename=f'email_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Load environment variables
load_dotenv()

def send_email(sender_email, sender_password, recipient_email, subject, body, smtp_server="smtp.gmail.com", smtp_port=587):
    """
    Send an email using SMTP
    """
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Add body
        msg.attach(MIMEText(body, 'plain'))

        # Create SMTP session
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        
        logging.info(f"Email sent successfully to {recipient_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {recipient_email}: {str(e)}")
        return False

def process_excel_data(excel_file):
    """
    Process Excel file and send emails
    """
    try:
        # Read Excel file
        df = pd.read_excel(excel_file)
        
        # Required columns
        required_columns = ['sender_email', 'sender_password', 'recipient_email', 'subject', 'body']
        
        # Check if all required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

        # Get SMTP settings from environment variables or use defaults
        smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        smtp_port = int(os.getenv('SMTP_PORT', '587'))

        # Process each row
        success_count = 0
        failure_count = 0
        
        for index, row in df.iterrows():
            success = send_email(
                sender_email=row['sender_email'],
                sender_password=row['sender_password'],
                recipient_email=row['recipient_email'],
                subject=row['subject'],
                body=row['body'],
                smtp_server=smtp_server,
                smtp_port=smtp_port
            )
            
            if success:
                success_count += 1
            else:
                failure_count += 1

        logging.info(f"Email sending completed. Success: {success_count}, Failures: {failure_count}")
        return success_count, failure_count

    except Exception as e:
        logging.error(f"Error processing Excel file: {str(e)}")
        raise

def main():
    try:
        # Get Excel file path from user
        excel_file = input("Enter the path to your Excel file: ")
        
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
        
        # Process the Excel file and send emails
        success_count, failure_count = process_excel_data(excel_file)
        
        print(f"\nEmail sending completed!")
        print(f"Successful emails: {success_count}")
        print(f"Failed emails: {failure_count}")
        print(f"Check the log file for detailed information.")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        logging.error(f"Main program error: {str(e)}")

if __name__ == "__main__":
    main() 