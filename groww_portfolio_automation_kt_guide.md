# <selection-tag>Groww Portfolio Automation - Knowledge Transfer & Documentation Guide</selection-tag>

## 1. Introduction and Purpose

This project is a fully automated data pipeline designed to track, format, and email a monthly Mutual Fund portfolio report.

**The Old Manual Process:** Previously, generating a monthly report required manually logging into Groww, downloading a raw Excel file, manually formatting the data, calculating total returns, splitting specific funds into separate sheets, calculating the overall Net Worth (including external deposits), and manually drafting an email to family members.
**The Automated Solution:** Now, the user only has to click "Download" on the Groww website. The background automation instantly detects the downloading file, waits for it to finish, extracts the exact date, automatically categorizes the funds, generates a highly formatted 3-sheet Excel workbook, archives the raw data, and emails the final report to the family—all within seconds, with zero manual intervention.

## 2. End-to-End Workflow

When the automation is running in the background, the pipeline follows this exact sequence:

1. **Trigger (Download):** A raw Excel report is downloaded from Groww into the `Downloads` folder.
2. **Detection:** The `watchdog` library detects a new file starting with `mutual_fund...`.
3. **Readiness Check:** The script continuously checks the file to ensure the web browser has completely finished downloading it before trying to read it.
4. **Extraction:** `Pandas` opens the raw Excel file and scans the text to find the exact "Holdings As On \[Date\]" to accurately date the report.
5. **Transformation:** The raw data is cleaned. Specific "Equity" funds are filtered out from the main "Liquid" funds.
6. **Excel Generation:** The script creates a new Excel workbook containing three formatted sheets: *Consolidated Folio*, *Liquid Funds*, and *Equity Funds*.
7. **Archiving:** The raw, unformatted file is moved out of the `Downloads` folder and into an `Archive` folder.
8. **Transmission:** The final formatted report is passed to the email module, attached to an HTML email, and sent to the recipients.
9. **Logging:** Every step (successes and errors) is silently recorded in a local log file.

## 3. Technology and Prerequisites

Before setting up this automation on a new machine, ensure you have the following:

* **Operating System:** Windows 10 or 11.
* **Python Version:** Python 3.9 or higher.
* **Installing Python:**
  1. Download the installer from [python.org](https://www.python.org/downloads/).
  2. **CRITICAL:** During installation, check the box that says **"Add Python to PATH"** before clicking Install.
* **Verification:** Open Command Prompt and type `python --version`. It should print your Python version.
* **Code Editor:** A basic text editor like Notepad, or an IDE like Visual Studio Code (recommended) to edit configurations.

**Email Prerequisites (Choose One Based on Your Environment):**
* **For Personal Laptops:** A standard Gmail account with 2-Step Verification enabled, allowing you to generate an **App Password**.
* **For Office/Corporate Laptops:** A free **Brevo** (formerly Sendinblue) account. You will need to generate a `v3 API Key` and verify your sender email address in their dashboard.

## 4. Python Package Installation

This automation relies on a few external Python libraries. Open Command Prompt or PowerShell and run the following command to install all of them at once:

```text
pip install pandas watchdog openpyxl xlsxwriter requests
```

**Why are these required?**
* `pandas`: The core data engine used to read the raw Excel file, clean columns, and calculate math/percentages.
* `watchdog`: The monitoring library that constantly watches the `Downloads` folder for new files.
* `openpyxl`: Required by Pandas to *read* `.xlsx` files.
* `xlsxwriter`: Required to *write* the highly formatted final Excel file.
* `requests`: Required to send HTTP POST requests if you are using the Brevo API for email on an office laptop.

## 5. Project Folder Structure

Ensure your files are organized securely in a dedicated project folder (e.g., `C:\Users\YourName\Desktop\Groww\Scripts`).

* **`config.py`**: The central brain for settings. Holds all folder paths, email addresses, and passwords/API keys.
* **`Scripts.py`**: The main engine. Contains the Watchdog monitor, the Pandas data transformation logic, and the Excel drawing logic.
* **`email_service.py`**: The dedicated module that constructs the HTML email and transmits it.
* **`Trigger.bat`**: A Windows batch file used to easily start the script invisibly in the background.
* **`automation.log`**: A text file (auto-generated) that records exactly what the script is doing in real-time.

You will also need three external folders (defined in `config.py`):
1. **Watch Folder:** Usually your standard `Downloads` folder.
2. **Output Folder:** Where the final, formatted reports are saved.
3. **Archive Folder:** Where the raw Groww downloads are moved to keep your Downloads folder clean.

## 6. Configuration Setup (`config.py`)

When moving this script to a new laptop, you **must** update `config.py`. Do not change paths inside the main script.

```python
# config.py
import os

# --- 1. Folder Paths (MUST BE UPDATED FOR NEW LAPTOPS) ---
WATCH_FOLDER = r"C:\Users\YOUR_USERNAME\Downloads"
ARCHIVE_FOLDER = r"C:\Users\YOUR_USERNAME\Desktop\Groww\Archive"
OUTPUT_FOLDER = r"C:\Users\YOUR_USERNAME\Desktop\Groww\Output"
LOG_FILE = r"C:\Users\YOUR_USERNAME\Desktop\Groww\Scripts\automation.log"

# --- 2. Email Credentials ---
SENDER_EMAIL = "your_email@gmail.com"

# Use this if running on a Personal Laptop (Standard SMTP)
GMAIL_APP_PASSWORD = "<YOUR_16_DIGIT_APP_PASSWORD>" 

# Use this if running on an Office Laptop (Corporate Firewall Bypass)
BREVO_API_KEY = "<YOUR_BREVO_V3_API_KEY>"

# --- 3. Recipients ---
RECEIVER_EMAILS = ["dad_email@example.com"] 
# RECEIVER_EMAILS can also be a list: ["email1@test.com", "email2@test.com"]
```

## 7. Report Processing Logic

* **Date Identification:** The script does *not* rely on the day you downloaded the file. It opens the Excel file and searches for the text `HOLDINGS AS ON [Date]`. This ensures that if you download a historical report from 3 months ago, it accurately processes and labels it with the old date.
* **Data Transformation:** It renames raw columns (e.g., "Invested Value" to "Invested_Value") to make math easier. It calculates "Returns" and "Percentage" for every row.
* **Fund Segregation:** The script looks at a hardcoded list of funds (e.g., *Quant Flexi Cap*). Any row matching those names is moved into an "Equity" data bucket, while the rest remain in the "Liquid" data bucket.
* **External Balances:** Static values for a Credit Card Deposit (18,000) and House Deposit (40,000) are added to the Liquid Funds total to calculate the Grand Total Net Worth.

## 8. Excel Report Structure

The final generated Excel file contains three highly formatted tabs:

1. **Consolidated Folio:** A master overview. Includes a "Category Allocation" table showing what percentage of the portfolio is Liquid vs. Equity. It lists *all* funds together to show combined Profit/Loss.
2. **Liquid Funds:** Tracks the main mutual funds. This is the only sheet that includes the final "Grand Total" Net Worth (including external deposits).
3. **Equity Funds:** A completely isolated view of specific, independently tracked equity funds.

**Formatting Rules Applied Automatically:**
* Long fund names will dynamically widen the column so text is never cut off.
* Profitable returns/percentages are colored **Green**. Loss-making returns are colored **Red**.
* Odd/Even rows have alternating white and grey backgrounds for readability.
* A "Weight" column is calculated to show how much of the portfolio a single fund consumes.

## 9. File Monitoring and Multiple Downloads

* **Detection:** `watchdog` listens for events in the `Downloads` folder. If a file ends in `.xlsx` and starts with `mutual_fund`, it triggers.
* **File Locks:** Browsers download files in chunks. If the script tries to read a half-downloaded file, it will crash. The script uses a `wait_for_file_unlock` function that gently tries to open the file in "append" mode. If Windows denies access, the script knows the browser is still downloading it, and waits patiently.
* **Multiple Simultaneous Downloads:** If you download 3 reports at the same time, you do not need to worry about race conditions or them overwriting each other. Watchdog places them in a **queue**. It will completely finish processing and emailing Report A before it even looks at Report B.

## 10. Email Automation (Personal vs. Office Setup)

Because corporate networks have strict firewalls, the codebase supports two distinct ways to send the email depending on which environment you are running it in. You must configure `email_service.py` to use the appropriate method.

### Method A: Gmail SMTP (For Personal Laptops)
* **How it works:** Logs into Gmail using `smtplib` on Port 465 (SSL). 
* **Requirements:** A Google App Password. Standard Google passwords will fail due to Google's security blocks.
* **Pros/Cons:** Easiest to set up, but frequently blocked by corporate VPNs/Firewalls.

### Method B: Brevo API (For Office/Corporate Laptops)
* **How it works:** Corporate firewalls often block outbound SMTP traffic (Ports 465/587). The Brevo API bypasses this by wrapping the email data into a standard HTTPS POST request (Port 443), which corporate networks allow because it looks like regular web browsing traffic.
* **Setup Instructions:**
  1. Create a free account at [brevo.com](https://www.brevo.com/).
  2. Verify your sender email address in their dashboard.
  3. Navigate to **SMTP & API** -> **API Keys** and generate a new `v3 API Key`.
  4. Paste this key into `config.py` under `BREVO_API_KEY`.
  5. The `email_service.py` module must be swapped to the `requests.post` version of the code that points to `https://api.brevo.com/v3/smtp/email`.

## 11. Logging

* All console output, `print()` statements, and hidden error tracebacks are safely caught and written to `automation.log` inside your Scripts folder.
* **Log Rotation:** The script uses a `RotatingFileHandler`. If the log file reaches 5 Megabytes, it backs it up and starts a fresh one. It keeps a maximum of 3 backups. This guarantees the automation will never slowly fill up your computer's hard drive over the years.
* **Troubleshooting:** If the script "isn't doing anything," open `automation.log`. Scroll to the very bottom to see the exact error message.

## 12. Running the Automation

* **To Start:** Double-click the `Trigger.bat` file. A black window will flash for a split second and disappear. The script is now running silently in the background.
* **To Verify:** Open Task Manager, go to the "Details" tab, and look for `pythonw.exe`.
* **To Stop:** You can end the `pythonw.exe` task in Task Manager, or open Command Prompt and run: `taskkill /F /IM pythonw.exe`.
* **Restarting the Computer:** If you turn off your laptop, the script dies. You must double-click `Trigger.bat` again when you turn the laptop back on. (Alternatively, place a shortcut to `Trigger.bat` in your Windows `shell:startup` folder to make it run automatically on boot).

## 13. Testing and Validation (For a New Setup)

When deploying to a new laptop, follow these steps to validate:

1. Open PowerShell in your Scripts folder.
2. Run the script directly in the terminal so you can see errors: `python .\Scripts.py`
3. Download a historical report from Groww.
4. **Expected Output:** The terminal should print `[TRIGGER] Groww report detected...`, followed by the report date, and finally `[SUCCESS] Email sent successfully...`.
5. Check your `Archive` folder (the raw file should be there).
6. Check your `Output` folder (the formatted file should be there).
7. Check your email inbox.
8. Stop the terminal script (`Ctrl + C`) and start it via the `.bat` file for background use.

## 14. Troubleshooting Guide

| **Symptom** | **Likely Cause** | **Diagnosis & Fix** | 
| **Command Prompt says "python is not recognized"** | Python isn't in your System PATH. | Reinstall Python and ensure "Add Python to PATH" is checked at the bottom of the installer. | 
| **Terminal throws `ModuleNotFoundError`** | Missing packages. | Run `pip install pandas watchdog openpyxl xlsxwriter requests` in your terminal. | 
| **I downloaded the file, but absolutely nothing happened.** | Folder path mismatch or Watchdog isn't running. | Check `config.py` to ensure `WATCH_FOLDER` points to your exact Downloads folder. Ensure the background script is actually running via Task Manager. | 
| **File processes, but email fails to send (SMTP Timeout)** | Corporate Firewall blocking Port 465. | Switch `email_service.py` to use the Brevo API integration instead of standard SMTP. | 
| **Brevo Error: "Your Brevo account has been suspended"** | Automated Anti-Spam Bot flag. | New free Brevo accounts sending `.xlsx` attachments to Gmail are often flagged as spam bots. Reply directly to the suspension email stating this is a personal Python script sending a financial report to family. They will manually lift the suspension. |
| **Log file is not updating.** | Logging redirect is broken. | Ensure `sys.stdout = StreamToLogger...` is NOT commented out in `Scripts.py`. | 
| **Error: "Could not find HOLDINGS AS ON"** | Groww changed their export format. | Open the raw Excel file. Did Groww change the header text? You may need to update the regex search string in `Scripts.py`. | 

## 15. Maintenance and Future Changes

* **Changing Laptops:** When migrating from a personal laptop to an office laptop, be prepared to swap `email_service.py` to the API version to bypass the corporate firewall.
* **Adding new Equity Funds:** Open `Scripts.py`. Locate the list named `SEPARATED_FUNDS`. Simply add the exact name of the new fund to this list (in quotes, separated by commas). The script will automatically route it to the Equity sheet next month.
* **Updating external deposits:** In `Scripts.py`, locate `deposit_cc = 18000` and `house_deposit = 40000`. Update these numbers as your real-world balances change.
* **DO NOT TOUCH:** Do not casually edit the `wait_for_file_unlock` function or the `watchdog` event classes. Modifying these can reintroduce the duplicate-processing bugs associated with browser downloads.

## 16. Security and Best Practices

* **Never hardcode passwords:** Always keep your App Password and Brevo API Key inside `config.py`. If you ever share `Scripts.py` with someone to show off your code, your passwords will remain safe.
* **API Key Management:** Treat a Google App Password and Brevo API key like a master key. Do not email them to yourself or save them in plain text anywhere other than the `config.py` file. If you suspect they leaked, go to the respective security settings and revoke them instantly.

## 17. Quick Reference / KT Cheat Sheet

* **Installation:** `pip install pandas watchdog openpyxl xlsxwriter requests`
* **Start Command (Foreground):** `python .\Scripts.py`
* **Start Command (Background):** Double click `Trigger.bat`
* **Stop Command:** `taskkill /F /IM pythonw.exe`
* **Config File:** `config.py` (Update paths and emails here)
* **Log Location:** Look for `automation.log` in your Scripts folder.
* **Email Methods:** Personal Laptop = Gmail SMTP. Office Laptop = Brevo API.
* **Core Logic:** Detect `.xlsx` -> Extract Date -> Split Funds -> Build 3 Sheets -> Send Email.