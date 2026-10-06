import os
import smtplib
from email.message import EmailMessage
import config

def send_report_email(excel_path, pdf_path, report_date):
    print(f"\nPreparing to send email via standard Gmail SMTP for: {report_date}")
    
    msg = EmailMessage()
    msg['Subject'] = f"Monthly Mutual Fund Balance Sheet - {report_date}"
    
    # Custom sender name
    msg['From'] = f"Groww Balancesheet <{config.SENDER_EMAIL}>"
    
    # Safe checks for To, Cc, and Bcc
    if config.RECEIVER_EMAILS:
        msg['To'] = ", ".join(config.RECEIVER_EMAILS) if isinstance(config.RECEIVER_EMAILS, list) else config.RECEIVER_EMAILS
            
    if config.CC_EMAILS:
        msg['Cc'] = ", ".join(config.CC_EMAILS) if isinstance(config.CC_EMAILS, list) else config.CC_EMAILS
            
    if config.BCC_EMAILS:
        msg['Bcc'] = ", ".join(config.BCC_EMAILS) if isinstance(config.BCC_EMAILS, list) else config.BCC_EMAILS

    html_content = f"""
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
        <p>Hi,</p>
        <p>Please find attached the updated <strong>Mutual Fund Balance Sheet & Net Worth Summary</strong> for the period ending <strong>{report_date}</strong>.</p>
        <p>Best regards,<br><strong>Pranav Shirali</strong></p>
        <hr style="border: none; border-top: 1px solid #eee; margin-top: 20px;">
        <p style="font-size: 12px; color: #888;">
            <em>Note: This email and the attached reports were generated automatically by a personal automation.</em>
        </p>
    </div>
    """
    msg.add_alternative(html_content, subtype='html')

    # Process Multiple Attachments
    files_to_attach = [excel_path, pdf_path]
    
    for file_path in files_to_attach:
        if not file_path or not os.path.exists(file_path):
            continue
            
        file_name = os.path.basename(file_path)
        try:
            with open(file_path, "rb") as f:
                file_data = f.read()
            
            # Determine correct MIME subtype (Excel vs PDF)
            subtype = "vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_name.endswith('.xlsx') else "pdf"
                
            msg.add_attachment(file_data, maintype="application", subtype=subtype, filename=file_name)
        except Exception as e:
            print(f"[ERROR] Failed to attach {file_name}: {e}")

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(config.SENDER_EMAIL, config.GMAIL_APP_PASSWORD)
            smtp.send_message(msg)
        print(f"[SUCCESS] Email sent successfully with {len(files_to_attach)} attachments.")
    except Exception as e:
        print(f"[ERROR] SMTP connection failed: {e}")