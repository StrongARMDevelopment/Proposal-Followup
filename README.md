Automated Proposal Follow-Up Script
Version: 2.2.0
Author: [Aaron Melton/Harbor Fab]
Last Updated: June 2025

Overview
This Python script automates the process of sending follow-up emails for submitted proposals. It reads data from Excel logs, identifies proposals that are due for a follow-up based on configurable rules, and sends personalized, consolidated emails to each contact.

The primary goal is to maintain professional communication with clients without overwhelming them with individual emails for each proposal. If a single contact has multiple proposals ready for follow-up, the script intelligently groups them into a single, well-formatted email.

This tool is designed to be robust, configurable, and safe, featuring multiple modes of operation including a "dry run" for testing and a comprehensive test email function.

Key Features
Consolidated Emails: Groups multiple proposals for the same contact into a single, clean email.

Multi-Stage Follow-Up Logic: Uses different email snippets based on the follow-up stage (1st, 2nd, 3rd+).

External Configuration: All settings, paths, and email templates are managed in an external config.ini file, so no code changes are needed for adjustments.

Safe Operation Modes:

Live Run: The default mode that sends emails and updates Excel files.

Dry Run: Simulates the entire process and logs what would happen without sending emails or modifying files.

Test Email Mode: Sends pre-formatted test emails to a specified recipient to verify template appearance.

Robust Error Handling: Includes specific checks for file access, network issues, and invalid data.

Direct Outlook Signature Integration: Reads your named Outlook signature directly from its file, ensuring emails are perfectly formatted.

Dynamic Column Lookup: Finds columns in Excel by header name, so reordering columns in your log files won't break the script.

Safety Features:

Creates automatic .bak backups of Excel files before modifying them.

Uses a .lock file to prevent multiple instances of the script from running simultaneously.

Detailed Logging: Logs all actions, warnings, and errors to both the console and a dedicated log file (followup_automation.log).

Requirements
Python 3.8+

Required Python libraries:

pandas

openpyxl

pywin32

zipfile

shutil

Setup & Configuration
1. Installation
Clone the repository and install the required packages using pip:

git clone [your-repository-url]
cd [your-repository-folder]
pip install -r requirements.txt

(Note: You will need to create a requirements.txt file containing pandas, openpyxl, and pywin32 or install them manually.)

pip install pandas openpyxl pywin32

2. Configuration (config.ini)
Before running the script, you must set up the config.ini file. This file controls every aspect of the script's behavior.

[Settings]
DryRun: True or False. Set to True to test logic without consequences.

SendTestEmail: True or False. If True, the script will only send test emails and will not process Excel files.

TestEmailRecipient: The email address where test emails should be sent.

SignatureName: The exact name of your signature in Outlook (e.g., Standard (amelton@harborfab.com)). The script will look for the corresponding .htm file.

YearsToProcess: Comma-separated list of years to process (e.g., 2024,2025).

DaysFirstFollowUp: Number of days after submission (or last contact) before the first follow-up.

DaysSubsequentFollowUps: Number of days between later follow-ups.

...and other settings for email retries, delays, and logging.

[Paths]
Log<Year>Path: The full file path to your Excel proposal logs for each year specified in YearsToProcess.

ScriptLogFile: Name of the file for detailed logging.

LockFilePath: Name of the lock file to prevent concurrent runs.

[Columns]
Maps the script's internal variable names to the exact header names used in your Excel files. This is crucial for the script to find the right data.

[EmailSubjects], [EmailBody], [EmailSnippets]
These sections control the content of the emails. You can customize the greeting, intro, closing, and the specific message for each follow-up stage using {placeholders} which the script will populate with data.

Usage
You can run the script in one of three modes, controlled by the config.ini file.

1. Dry Run Mode (Recommended for Testing)
This is the safest way to test. It will read your Excel files and generate logs of which emails it would send and which rows it would update, without changing anything.

Set DryRun = True in config.ini.

Set SendTestEmail = False.

Run the script: python follow_up.py

Review followup_automation.log to see the simulated actions.

2. Test Email Mode
This mode is for checking the look and feel of your email templates. It does not read your Excel files.

Set SendTestEmail = True in config.ini.

Ensure TestEmailRecipient is set to your email address.

Run the script: python follow_up.py

You will receive three separate test emails showcasing the templates for a single, double, and multi-project follow-up.

3. Live Run
This is the main operational mode. It will send emails to customers and update your Excel logs.

Backup your Excel files! (Though the script creates .bak files, manual backups are always wise).

Set DryRun = False in config.ini.

Set SendTestEmail = False.

Run the script: python follow_up.py

How It Works
The script operates in a clear, sequential process:

Initialization: Loads settings from config.ini, sets up logging, and creates a lock file.

Phase 1: Collect Data: It iterates through the specified Excel files and sheets, reading each row and applying the follow-up logic. Any proposal that meets the criteria is added to a collection, grouped by the contact's email address.

Phase 2: Send Emails: It loops through the collection of grouped follow-ups. For each contact, it constructs a single, consolidated email using the templates and snippets from config.ini, then sends it via Outlook.

Phase 3: Update Excel: After successfully sending emails, it opens the Excel files and updates the "Last Correspondence" and "Follow-Up Stage" columns for all proposals that were included in the emails.

Cleanup: Removes the lock file upon completion or if an error occurs.
* Waits between emails to avoid Outlook rate limits (see `EmailSendDelaySeconds`).
* Dry-run mode (`DryRun = True`) allows you to test without sending emails or modifying Excel files.
* Excel files are backed up before changes if `BackupExcelBeforeSave = True`.

---
