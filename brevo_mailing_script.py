import os
import base64
import logging
import requests

import config

# Connect to the central logger we built in Scripts.py
logger = logging.getLogger(__name__)

def send_report_email(file_path, report_date):
    logger.info(f"Preparing to send email via Brevo API for: {os.path.basename(file_path)}")
    
    if not os.path.exists(file_path):
        logger.error(f"[ERROR] Cannot send email. File not found: {file_path}")
        return

    file_name = os.path.basename(file_path)

    # 1. Base64-encode the specific file passed by Scripts.py
    try:
        with open(file_path, "rb") as f:
            encoded_file = base64.b64encode(f.read()).decode('utf-8')
    except Exception as e:
        logger.error(f"[ERROR] Failed to read the Excel file: {e}")
        return

    # 2. Construct API Request
    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "accept": "application/json",
        "api-key": config.BREVO_API_KEY,
        "content-type": "application/json"
    }
    
    payload = {
        "sender": {"email": config.SENDER_EMAIL, "name": "pranavshirali.work"},
        "to": [{"email": email} for email in config.RECEIVER_EMAILS],
        "subject": f"Monthly Mutual Fund Balance Sheet - {report_date}",
        "htmlContent": f"""
        <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
            <p>Hi,</p>
            <p>Please find attached the updated <strong>Mutual Fund Balance Sheet & Net Worth Summary</strong> for the period ending <strong>{report_date}</strong>.</p>
            <p>Best regards,<br><strong>Pranav Shirali</strong></p>
            <hr style="border: none; border-top: 1px solid #eee; margin-top: 20px;">
            <p style="font-size: 12px; color: #888;">
                <em>Note: This email and the attached report were generated automatically by a personal automation.</em>
            </p>
        </div>
        """,
        "attachment": [{"content": encoded_file, "name": file_name}]
    }
    
    if config.CC_EMAILS:
        payload["cc"] = [{"email": email} for email in config.CC_EMAILS]
    if config.BCC_EMAILS:
        payload["bcc"] = [{"email": email} for email in config.BCC_EMAILS]

    # 3. Transmit
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=15)
        if response.status_code in (201, 202):
            logger.info(f"[SUCCESS] Email sent successfully with attachment: {file_name}")
        else:
            logger.error(f"[ERROR] API rejected the request. Code: {response.status_code} | Details: {response.text}")
    except Exception as e:
        logger.error(f"[ERROR] HTTPS connection failed during transmission: {e}")