import os
import smtplib
from email.message import EmailMessage
import config

def send_report_email(file_path, report_date):
    print(f"\nPreparing to send email via standard Gmail SMTP for: {os.path.basename(file_path)}")
    
    if not os.path.exists(file_path):
        print(f"[ERROR] Cannot send email. File not found: {file_path}")
        return

    file_name = os.path.basename(file_path)

    msg = EmailMessage()
    msg['Subject'] = f"Monthly Mutual Fund Balance Sheet - {report_date}"
    msg['From'] = config.SENDER_EMAIL
    
    if isinstance(config.RECEIVER_EMAILS, list):
        msg['To'] = ", ".join(config.RECEIVER_EMAILS)
    else:
        msg['To'] = config.RECEIVER_EMAILS
    
    if isinstance(config.CC_EMAILS, list):
        msg['Cc'] = ", ".join(config.CC_EMAILS)
    else:
        msg['Cc'] = config.CC_EMAILS
    
    if isinstance(config.BCC_EMAILS, list):
        msg['Bcc'] = ", ".join(config.BCC_EMAILS)
    else:
        msg['Bcc'] = config.BCC_EMAILS

    html_content = f"""
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
        <p>Hi,</p>
        <p>Please find attached the updated <strong>Mutual Fund Balance Sheet & Net Worth Summary</strong> for the period ending <strong>{report_date}</strong>.</p>
        <p>Best regards,<br><strong>Pranav Shirali</strong></p>
        <hr style="border: none; border-top: 1px solid #eee; margin-top: 20px;">
        <p style="font-size: 12px; color: #888;">
            <em>Note: This email and the attached report were generated automatically by a personal automation.</em>
        </p>
    </div>
    """
    msg.add_alternative(html_content, subtype='html')

    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
        
        msg.add_attachment(
            file_data, 
            maintype="application", 
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            filename=file_name
        )
    except Exception as e:
        print(f"[ERROR] Failed to attach the Excel file: {e}")
        return

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            # Requires config.GMAIL_APP_PASSWORD to be defined in config.py
            smtp.login(config.SENDER_EMAIL, config.GMAIL_APP_PASSWORD)
            smtp.send_message(msg)
        print(f"[SUCCESS] Email sent successfully with attachment: {file_name}")
    except Exception as e:
        print(f"[ERROR] SMTP connection failed: {e}")