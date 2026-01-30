import pandas as pd
import smtplib
import time
import ssl
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# =======================================================
# ⚙️ CONFIGURATION (ADJUST SETTINGS HERE)
# =======================================================

# 1. Sender and Server Credentials
SENDER_EMAIL = "your-email@domain.com"
SENDER_PASSWORD = "your-app-password"  # Recommendation: Use an 'App Password'
SMTP_SERVER = "smtppro.zoho.com"
SMTP_PORT = 465  # Use 465 for SSL or 587 for TLS/STARTTLS

# 2. Optional Recipient Settings
# To enable CC, remove the '#' below and enter the email address
# CC_EMAIL = "manager@domain.com"

SUBJECT = " [SUBJECT HERE] "
EXCEL_FILE = "mail.xlsx"  # Name of your Excel file

# 3. Anti-Spam Timing Settings (Seconds)
MIN_WAIT = 45        # Minimum delay between emails
MAX_WAIT = 90        # Maximum delay between emails
COOLDOWN_EVERY = 10  # Number of emails before a long break
COOLDOWN_TIME = 300  # Duration of the long break (5 minutes)

# 4. Email Signature (HTML Format)
EMAIL_SIGNATURE = """
<br><br>
<div style="font-family: Arial, Helvetica, sans-serif; font-size: 10pt; line-height: 1.5;">
    <b>Best Regards,</b>
    <br><br>
    <b>John Doe</b>
    <br><br>
    <b>Business Development Manager</b>
    <br>
    <b>Turn Your Ideas into Products...</b>
    <br>
    <a href="http://company.com" target="_blank">company.com</a>
</div>
"""

# =======================================================
# 🛠️ EMAIL SENDER FUNCTION
# =======================================================

def send_email(to_email, body_text):
    try:
        # Convert newlines to HTML breaks and append signature
        html_body = body_text.replace('\n', '<br>') + EMAIL_SIGNATURE
        
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = SUBJECT
        
        # Add CC if the variable is defined in the configuration
        if 'CC_EMAIL' in globals():
            msg['Cc'] = CC_EMAIL

        msg.attach(MIMEText(html_body, 'html'))
        context = ssl.create_default_context()
        
        # Connection logic for SSL (465) or STARTTLS (587)
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context, timeout=30) as server:
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        else:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
                server.starttls(context=context)
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        return "SUCCESS"

    except smtplib.SMTPResponseException as e:
        if e.smtp_code == 550:
            return "SPAM_BLOCK"
        return f"SMTP_ERROR_{e.smtp_code}: {str(e.smtp_error)}"
    except Exception as e:
        return f"CONNECTION_ISSUE: {str(e)}"

# =======================================================
# 🚀 MAIN EXECUTION LOOP
# =======================================================

def main():
    print(f"--- 📬 Mailer Starting as {SENDER_EMAIL} ---")
    
    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"🔴 ERROR: Could not read Excel file: {e}")
        return

    sent_count = 0
    for index, row in df.iterrows():
        # Clean and split multiple emails if they exist in the cell
        raw_emails = str(row['mail'])
        body_text = row['message']
        email_list = [email.strip() for email in raw_emails.replace(',', ';').split(';') if email.strip()]
        
        for email_addr in email_list:
            print(f"📧 Processing: {email_addr}...")
            result = send_email(email_addr, body_text)
            
            if result == "SUCCESS":
                sent_count += 1
                wait_time = random.randint(MIN_WAIT, MAX_WAIT)
                print(f"✅ Sent! Waiting {wait_time}s to avoid spam filters...")
                
                # Check for cooldown period
                if sent_count % COOLDOWN_EVERY == 0:
                    print(f"☕ {COOLDOWN_EVERY} emails sent. Starting {COOLDOWN_TIME}s cooldown...")
                    time.sleep(COOLDOWN_TIME)
                else:
                    time.sleep(wait_time)

            elif result == "SPAM_BLOCK":
                print("🛑 SPAM BLOCK DETECTED: Stopping script for security.")
                return 
            
            else:
                print(f"❌ {result}")
                print("⚠️ Retrying next address in 15s...")
                time.sleep(15)

if __name__ == "__main__":
    main()