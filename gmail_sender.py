import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
import logging
from datetime import datetime
import openpyxl
import sys
import json
import base64
import csv
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

# Set up logging
logging.basicConfig(
    filename=f'gmail_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

CONFIG_FILE = 'gmail_config.json'
SALT = b'gmail_sender_salt'  # This should be stored securely in production

def print_app_password_instructions():
    """Print instructions for getting an App Password"""
    print("\n=== Gmail App Password Instructions ===")
    print("1. Go to your Google Account settings (https://myaccount.google.com/)")
    print("2. Click on 'Security'")
    print("3. Under '2-Step Verification', click on 'App passwords'")
    print("4. Select 'Mail' as the app and 'Other' as the device")
    print("5. Click 'Generate'")
    print("6. Copy the 16-character password (it will look like: xxxx xxxx xxxx xxxx)")
    print("7. Use this password when prompted for 'Gmail App Password'")
    print("=======================================\n")

def generate_key(password):
    """Generate encryption key from password"""
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=SALT,
        iterations=100000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    return key

def save_credentials(email, password, master_password):
    """Save encrypted credentials to config file"""
    try:
        # Generate encryption key from master password
        key = generate_key(master_password)
        f = Fernet(key)
        
        # Encrypt the credentials
        credentials = {
            'email': email,
            'password': password
        }
        encrypted_data = f.encrypt(json.dumps(credentials).encode())
        
        # Save to file
        with open(CONFIG_FILE, 'wb') as file:
            file.write(encrypted_data)
        
        print("\nCredentials saved successfully!")
        return True
    except Exception as e:
        logging.error(f"Error saving credentials: {str(e)}")
        return False

def load_credentials(master_password):
    """Load and decrypt credentials from config file"""
    try:
        if not os.path.exists(CONFIG_FILE):
            return None, None
        
        # Generate encryption key from master password
        key = generate_key(master_password)
        f = Fernet(key)
        
        # Read and decrypt the file
        with open(CONFIG_FILE, 'rb') as file:
            encrypted_data = file.read()
        
        decrypted_data = f.decrypt(encrypted_data)
        credentials = json.loads(decrypted_data)
        
        return credentials['email'], credentials['password']
    except Exception as e:
        logging.error(f"Error loading credentials: {str(e)}")
        return None, None

def create_sample_file():
    """
    Create a sample CSV file with the correct format
    """
    try:
        sample_file = 'sample_emails.csv'
        
        # Create sample data
        sample_data = [
            ['recipient_email', 'subject', 'body'],
            ['example1@gmail.com', 'Test Subject 1', 'This is test email 1'],
            ['example2@gmail.com', 'Test Subject 2', 'This is test email 2']
        ]
        
        # Write to CSV file
        with open(sample_file, 'w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(sample_data)
        
        print(f"\nCreated sample CSV file: {sample_file}")
        return sample_file
    except Exception as e:
        logging.error(f"Error creating sample file: {str(e)}")
        raise

def send_gmail(sender_email, app_password, recipient_email, subject, body):
    """
    Send an email using Gmail SMTP
    """
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Add body
        msg.attach(MIMEText(body, 'plain'))

        # Gmail SMTP settings
        smtp_server = "smtp.gmail.com"
        smtp_port = 587

        # Create SMTP session
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            try:
                server.login(sender_email, app_password)
            except smtplib.SMTPAuthenticationError:
                print("\nError: Invalid App Password. Please make sure you're using an App Password, not your regular Gmail password.")
                print("Follow the instructions above to generate an App Password.")
                raise
            server.send_message(msg)
        
        logging.info(f"Email sent successfully to {recipient_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {recipient_email}: {str(e)}")
        return False

def process_file(file_path, sender_email, app_password):
    """
    Process CSV or Excel file and send emails using Gmail
    """
    try:
        print(f"\nReading file: {file_path}")
        
        # Determine file type and read accordingly
        if file_path.endswith('.csv'):
            # Read CSV file
            with open(file_path, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                headers = next(reader)  # Get headers
                
                # Validate headers
                required_columns = ['recipient_email', 'subject', 'body']
                missing_columns = [col for col in required_columns if col not in headers]
                if missing_columns:
                    raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
                
                # Process rows
                success_count = 0
                failure_count = 0
                
                for row in reader:
                    if not row or not row[0]:  # Skip empty rows
                        continue
                    
                    recipient_email, subject, body = row
                    
                    # Validate email
                    if not '@' in str(recipient_email):
                        logging.error(f"Invalid email format: {recipient_email}")
                        failure_count += 1
                        continue
                    
                    success = send_gmail(
                        sender_email=sender_email,
                        app_password=app_password,
                        recipient_email=recipient_email,
                        subject=subject,
                        body=body
                    )
                    
                    if success:
                        success_count += 1
                    else:
                        failure_count += 1
                
        else:
            # Try to read as Excel file
            try:
                df = pd.read_excel(file_path)
                
                # Validate headers
                required_columns = ['recipient_email', 'subject', 'body']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
                
                # Process rows
                success_count = 0
                failure_count = 0
                
                for _, row in df.iterrows():
                    if pd.isna(row['recipient_email']):  # Skip empty rows
                        continue
                    
                    # Validate email
                    if not '@' in str(row['recipient_email']):
                        logging.error(f"Invalid email format: {row['recipient_email']}")
                        failure_count += 1
                        continue
                    
                    success = send_gmail(
                        sender_email=sender_email,
                        app_password=app_password,
                        recipient_email=row['recipient_email'],
                        subject=row['subject'],
                        body=row['body']
                    )
                    
                    if success:
                        success_count += 1
                    else:
                        failure_count += 1
                        
            except Exception as e:
                logging.error(f"Error reading Excel file: {str(e)}")
                raise
            
        logging.info(f"Email sending completed. Success: {success_count}, Failures: {failure_count}")
        return success_count, failure_count

    except Exception as e:
        logging.error(f"Error processing file: {str(e)}")
        raise

def main():
    try:
        print("Gmail Bulk Email Sender")
        print("=======================")
        
        # Print App Password instructions
        print_app_password_instructions()
        
        # Check for saved credentials
        sender_email = None
        app_password = None
        
        if os.path.exists(CONFIG_FILE):
            print("\nFound saved credentials.")
            master_password = input("Enter your master password to load credentials (or press Enter to use new credentials): ").strip()
            
            if master_password:
                sender_email, app_password = load_credentials(master_password)
                if not sender_email or not app_password:
                    print("Failed to load saved credentials. Please enter new credentials.")
                    sender_email = None
                    app_password = None
        
        # If no saved credentials or failed to load, get new ones
        if not sender_email or not app_password:
            sender_email = input("Enter your Gmail address: ").strip()
            app_password = input("Enter your Gmail App Password (16-character password from Google Account): ").strip()
            
            # Ask if user wants to save credentials
            save = input("\nDo you want to save these credentials for future use? (y/n): ").strip().lower()
            if save == 'y':
                master_password = input("Enter a master password to encrypt your credentials: ").strip()
                if save_credentials(sender_email, app_password, master_password):
                    print("Credentials saved successfully!")
                else:
                    print("Failed to save credentials. Continuing without saving.")
        
        # Get file path
        file_path = input("\nEnter the path to your CSV or Excel file (or press Enter to create a sample file): ").strip()
        
        if not file_path:
            file_path = create_sample_file()
        elif not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Process the file and send emails
        print("\nSending emails...")
        success_count, failure_count = process_file(file_path, sender_email, app_password)
        
        print(f"\nEmail sending completed!")
        print(f"Successful emails: {success_count}")
        print(f"Failed emails: {failure_count}")
        print(f"Check the log file for detailed information.")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        logging.error(f"Main program error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 