import os
import base64
import logging
import requests
import config

logger = logging.getLogger(__name__)

def send_report_email(excel_path, pdf_path, report_date):
    logger.info(f"Preparing to send email via Brevo API for: {report_date}")
    
    # Process Multiple Attachments
    attachments = []
    for file_path in [excel_path, pdf_path]:
        if file_path and os.path.exists(file_path):
            try:
                with open(file_path, "rb") as f:
                    encoded_file = base64.b64encode(f.read()).decode('utf-8')
                attachments.append({"content": encoded_file, "name": os.path.basename(file_path)})
            except Exception as e:
                logger.error(f"[ERROR] Failed to read attachment {os.path.basename(file_path)}: {e}")

    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "accept": "application/json",
        "api-key": config.BREVO_API_KEY,
        "content-type": "application/json"
    }
    
    payload = {
        "sender": {"email": config.SENDER_EMAIL, "name": "Groww Balancesheet"},
        "to": [{"email": email} for email in config.RECEIVER_EMAILS],
        "subject": f"Monthly Mutual Fund Balance Sheet - {report_date}",
        "htmlContent": f"""
        <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
            <p>Hi,</p>
            <p>Please find attached the updated <strong>Mutual Fund Balance Sheet & Net Worth Summary</strong> for the period ending <strong>{report_date}</strong>.</p>
            <p>Best regards,<br><strong>Pranav Shirali</strong></p>
            <hr style="border: none; border-top: 1px solid #eee; margin-top: 20px;">
            <p style="font-size: 12px; color: #888;">
                <em>Note: This email and the attached reports were generated automatically by a personal automation.</em>
            </p>
        </div>
        """,
        "attachment": attachments
    }
    
    if config.CC_EMAILS:
        payload["cc"] = [{"email": email} for email in config.CC_EMAILS]
    if config.BCC_EMAILS:
        payload["bcc"] = [{"email": email} for email in config.BCC_EMAILS]

    try:
        response = requests.post(url, json=payload, headers=headers, timeout=15)
        if response.status_code in (201, 202):
            logger.info(f"[SUCCESS] Email sent successfully with {len(attachments)} attachments.")
        else:
            logger.error(f"[ERROR] API rejected the request. Code: {response.status_code} | Details: {response.text}")
    except Exception as e:
        logger.error(f"[ERROR] HTTPS connection failed during transmission: {e}")