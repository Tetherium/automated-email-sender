# Python Safe Bulk Mailer 📬

A robust and simple Python-based email automation tool designed to send personalized emails from an Excel database. This script includes built-in **anti-spam mechanisms** like randomized delays and cooldown periods to protect your sender reputation.

## 🚀 Features
* **Anti-Spam Logic:** Randomized wait times (45-90s) between emails.
* **Cooldown Periods:** Automatically pauses for 5 minutes after every 10 emails.
* **Excel Integration:** Reads recipient emails and custom messages directly from `.xlsx` files.
* **Multi-Recipient Support:** Handles multiple email addresses in a single Excel cell (separated by `;` or `,`).
* **HTML Support:** Send rich-text emails with professional signatures.
* **Dual Protocol:** Supports both SSL (Port 465) and STARTTLS (Port 587).

## 🛠️ Prerequisites

Before running the script, ensure you have:
1.  **Python 3.x** installed.
2.  An **App Password** for your email provider (especially for Zoho, Gmail, or Outlook).
3.  Your recipient list in an Excel file (`mail_test.xlsx`).

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/Tetherium/automated-email-sender.git
    cd automated-email-sender
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## 📋 Configuration

Open the script and update the **CONFIGURATION** section at the top:

* `SENDER_EMAIL`: Your email address.
* `SENDER_PASSWORD`: Your generated App Password.
* `SMTP_SERVER`: (Default is Zoho) Change if using another provider.
* `CC_EMAIL`: Uncomment this line if you want to keep someone in the loop.
* `EXCEL_FILE`: Ensure your Excel file has columns named `mail` and `message`.

## 🖥️ Usage

1.  Prepare your `mail_test.xlsx` file.
2.  Run the script:
    ```bash
    python mailer.py
    ```

## ⚠️ Important Note
This tool is intended for professional communication and legitimate business outreach. Do not use this script for sending unsolicited spam, as it may result in your email account being suspended by your provider.

## 📄 License
Distributed under the MIT License. See `LICENSE` for more information.

