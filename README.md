# Automated Proposal Follow-Up Script

**Version:** 2.2.0
**Author:** [Aaron Melton/Harbor Fab]
**Last Updated:** June 2025

---

## Overview

This Python script automates the process of sending follow-up emails for submitted proposals. It reads data from Excel logs, identifies proposals that are due for a follow-up based on configurable rules, and sends personalized, consolidated emails to each contact.

The primary goal is to maintain professional communication with clients without overwhelming them with individual emails for each proposal. If a single contact has multiple proposals ready for follow-up, the script intelligently groups them into a single, well-formatted email.

This tool is designed to be robust, configurable, and safe, featuring multiple modes of operation including a "dry run" for testing and a comprehensive test email function.

---

## Key Features

* **Consolidated Emails:** Groups multiple proposals for the same contact into a single, clean email.
* **Multi-Stage Follow-Up Logic:** Uses different email snippets based on the follow-up stage (1st, 2nd, 3rd+).
* **External Configuration:** All settings, paths, and email templates are managed in an external `config.ini` file, so no code changes are needed for adjustments.
* **Safe Operation Modes:**
    * **Live Run:** The default mode that sends emails and updates Excel files.
    * **Dry Run:** Simulates the entire process and logs what *would* happen without sending emails or modifying files.
    * **Test Email Mode:** Sends pre-formatted test emails to a specified recipient to verify template appearance.
* **Robust Error Handling:** Includes specific checks for file access, network issues, and invalid data.
* **Direct Outlook Signature Integration:** Reads your named Outlook signature directly from its file, ensuring emails are perfectly formatted.
* **Dynamic Column Lookup:** Finds columns in Excel by header name, so reordering columns in your log files won't break the script.
* **Safety Features:**
    * Creates automatic `.bak` backups of Excel files before modifying them.
    * Uses a `.lock` file to prevent multiple instances of the script from running simultaneously.
* **Detailed Logging:** Logs all actions, warnings, and errors to both the console and a dedicated log file (`followup_automation.log`).

---

## Requirements

* Python 3.8+
* Required Python libraries:
    * `pandas`
    * `openpyxl`
    * `pywin32`

*(Note: The `zipfile` and `shutil` libraries are part of the Python standard library and do not require separate installation.)*

---

## Setup & Configuration

### 1. Installation

Clone the repository and install the required packages using pip. It's recommended to do this within a virtual environment.

```bash
# Clone the repository (replace with your actual URL)
git clone [your-repository-url]
cd [your-repository-folder]

# Create and activate a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate  # On Windows, use `venv\Scripts\activate`

# Install required packages
pip install pandas openpyxl pywin32
