Proposal Follow-Up Automation

This Python script automates follow-up emails for sales proposals stored in annual Excel workbooks. It integrates with Microsoft Outlook (MAPI) to send personalized follow-ups based on proposal age and prior correspondence.

---

Features

* Reads from yearly Excel files (`2024` and `2025`) containing monthly sheets of proposal data.
* Sends follow-up emails using predefined templates and your Outlook signature.
* Skips projects marked as **Won**, **Lost**, or **Re-Bid**.
* Tracks and updates **Last Correspondence** and **Follow-Up Stage** columns in-place.
* Logs all actions and errors to `followup_errors.log`.

---

Requirements

* Windows OS
* Microsoft Outlook (configured and open)
* Python 3.8+
* Dependencies:

  * `pandas`
  * `openpyxl`
  * `pyxlsb`
  * `pywin32`

Install dependencies via:

```bash
pip install pandas openpyxl pyxlsb pywin32
```

---

Folder Structure

Update the following file paths as needed:

```python
file_paths = {
    "2025": r"H:\...\Proposals Submitted Log - 2025.xlsx",
    "2024": r"H:\...\Proposals Submitted Log - 2024.xlsx"
}
```

---

How It Works

1. **Setup**: Opens Outlook and retrieves the user's signature.
2. **Load Sheets**: Only processes sheets named after months.
3. **Row Checks**:

   * Skips rows with invalid or missing data.
   * Skips proposals marked won/lost/re-bid.
4. **Follow-Up Logic**:

   * Stage 0: 7+ days since submission.
   * Stage 1/2: 14+ days since last correspondence.
   * Stage 3+: Continues every 14 days.
5. **Email**: Selects a template based on the stage, fills in proposal info, and sends it via Outlook.
6. **Workbook Update**: Writes back updated dates and stages.

---

Running the Script

Open a terminal and run:

```bash
python your_script_name.py
```

Ensure Outlook is running and the Excel files are not open in another program.

---

Logging

All errors and actions are saved in `followup_errors.log` in the same directory as the script.

---

Customization

To change email templates, edit the `email_templates` list:

```python
email_templates = [
    (subject, html_body),
    ...
]
```

---

Notes

* Excel indexing assumes data starts on row 2 (after headers).
* Signature is extracted from a temporary draft email.
* Waits 1 second between emails to avoid Outlook rate limits.

---