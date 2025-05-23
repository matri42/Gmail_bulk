# Bulk Email Sender

This Python script allows you to send bulk emails using data from an Excel file.

## Setup

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Create a `.env` file (optional) with your SMTP settings:
```
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

## Excel File Format

Your Excel file should contain the following columns:
- `sender_email`: The email address to send from
- `sender_password`: The password for the sender email
- `recipient_email`: The email address to send to
- `subject`: The email subject
- `body`: The email body content

## Usage

1. Prepare your Excel file with the required columns (see sample_data.xlsx for reference)
2. Run the script:
```bash
python bulk_email_sender.py
```
3. Enter the path to your Excel file when prompted
4. The script will process the Excel file and send emails
5. Check the generated log file for detailed information about the email sending process

## Notes

- For Gmail users, you'll need to use an App Password instead of your regular password
- The script creates a log file with the format `email_log_YYYYMMDD_HHMMSS.log`
- Make sure your Excel file is properly formatted with all required columns
- The script includes error handling and will report any issues during the email sending process

## Security

- Never share your `.env` file or Excel file containing email passwords
- Consider using environment variables for sensitive information
- Use App Passwords for Gmail accounts instead of regular passwords 