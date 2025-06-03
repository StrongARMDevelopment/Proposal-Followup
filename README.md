# Proposal Follow-Up Automation

This Python script automates follow-up emails for sales proposals stored in annual Excel workbooks. It integrates with Microsoft Outlook (MAPI) to send personalized follow-ups based on proposal age and prior correspondence.

---

## Features

* Reads from yearly Excel files (e.g., `2024`, `2025`) containing monthly sheets of proposal data.
* Sends follow-up emails using predefined templates and your Outlook signature.
* Skips projects marked as **Won**, **Lost**, or **Re-Bid**.
* Tracks and updates **Last Correspondence** and **Follow-Up Stage** columns in-place.
* Logs all actions and errors to a configurable log file.
* Prevents concurrent runs with a lock file.
* Supports dry-run mode for safe testing.
* Backs up Excel files before making changes (optional, see config).

---

## Requirements

* Windows OS
* Microsoft Outlook (configured and open)
* Python 3.8+
* Dependencies:
  * `pandas`
  * `openpyxl`
  * `pywin32`

Install dependencies via:

```bash
pip install pandas openpyxl pywin32
```

---

## Configuration

All settings are managed in `config.ini`.  
**Example:**

```ini
[Settings]
DryRun = False
YearsToProcess = 2024,2025
ValidMonths = January,February,March,April,May,June,July,August,September,October,November,December
DaysFirstFollowUp = 7
DaysSubsequentFollowUps = 5
MaxEmailSendAttempts = 3
EmailRetryDelaySeconds = 5
EmailSendDelaySeconds = 1
DesiredOutlookAccount = 
LogLevel = INFO
BackupExcelBeforeSave = True

[Paths]
Log2024Path = H:\...\Proposals Submitted Log - 2024.xlsx
Log2025Path = H:\...\Proposals Submitted Log - 2025.xlsx
ScriptLogFile = followup_automation.log
LockFilePath = followup_automation.lock

[Columns]
DateProposalSubmitted = Date Submitted
LastCorrespondence = Last Correspondence
ContactEmail = Contact Email
ContactName = Contact Name
ProjectName = Project Name
Value = Value
Won = Won
Lost = Lost
ReBid = Re-Bid
FollowUpStage = Follow Up Stage

[EmailBody]
Greeting = <p>Hi {contact_first_name},</p>
Intro = <p>Just following up on the following proposal(s):</p>
ProjectListStart = <ul>
ProjectListItem = <li><strong>{project_name}</strong> ({submission_date_str}{value_info})<br/><em>{snippet_text}</em></li>
ProjectListEnd = </ul>
Closing = <p>Let us know if we can help.</p>
ValueFormat = , value: ${value_str}
DefaultSignatureFallback = <br><br>Best regards,

[EmailSubjects]
SingleProject = Follow-up on {project_name}
TwoProjects = Follow-up on {project_name_1} and {project_name_2}
MultipleProjects = Follow-up on {first_project_name} & others

[EmailSnippets]
Snippet_0 = We wanted to check in regarding this proposal.
Snippet_1 = Just a quick follow-up on our previous message.
Snippet_2 = This is a final courtesy follow-up regarding your proposal.
```

**Adjust file paths, column names, and templates as needed.**

---

## How It Works

1. **Setup**: Opens Outlook and retrieves the user's signature.
2. **Load Sheets**: Only processes sheets named after months listed in `ValidMonths`.
3. **Row Checks**:
   * Skips rows with invalid or missing data.
   * Skips proposals marked won/lost/re-bid.
4. **Follow-Up Logic**:
   * Stage 0: `DaysFirstFollowUp` days since submission or last correspondence.
   * Stage 1+: `DaysSubsequentFollowUps` days since last correspondence.
5. **Email**: Selects a template based on the stage, fills in proposal info, and sends it via Outlook.
6. **Workbook Update**: Writes back updated dates and stages. Optionally creates a backup before saving.
7. **Logging**: All actions and errors are logged to the file specified in `ScriptLogFile`.

---

## Running the Script

Open a terminal and run:

```bash
python follow_up.py
```

Ensure Outlook is running and the Excel files are not open in another program.

---

## Logging

All errors and actions are saved in the log file specified by `ScriptLogFile` in your `config.ini`.

---

## Customization

To change email templates, subjects, or snippets, edit the relevant sections in `config.ini`.

---

## Notes

* Excel indexing assumes data starts on row 2 (after headers).
* Signature is extracted from a temporary draft email.
* Waits between emails to avoid Outlook rate limits (see `EmailSendDelaySeconds`).
* Dry-run mode (`DryRun = True`) allows you to test without sending emails or modifying Excel files.
* Excel files are backed up before changes if `BackupExcelBeforeSave = True`.

---