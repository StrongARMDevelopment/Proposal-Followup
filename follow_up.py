import os
import pandas as pd
import datetime
import logging
import pyxlsb  # Required for reading .xlsb files
import win32com.client as win32  # Using MAPI instead of direct Outlook automation
from openpyxl import load_workbook  # Preserve formatting when updating Excel
import time  # Add time module for delay

# Configure logging
logging.basicConfig(filename="followup_errors.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Define file paths for both spreadsheets
file_paths = {
    "2025": r"H:\3 - Quotes\7 - Proposals Submitted Logs\Proposals Submitted Log - 2025.xlsx",
    "2024": r"H:\3 - Quotes\7 - Proposals Submitted Logs\Proposals Submitted Log - 2024.xlsx"
}

# Get today's date
today = datetime.datetime.today().date()

# Initialize Outlook using MAPI
try:
    outlook = win32.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    account = namespace.Accounts[0]
except Exception as e:
    logging.critical("Outlook is not available or not configured properly.")
    raise

# Retrieve the default Outlook signature
dummy_mail = outlook.CreateItem(0)
try:
    dummy_mail.Display()  # Opens a draft email to access signature
    signature = dummy_mail.HTMLBody  # Extracts the signature
    dummy_mail.Close(0)  # Closes the draft email without saving
except Exception as e:
    logging.error(f"Could not retrieve Outlook signature: {str(e)}")
    signature = ""

# Loop through each file path
for year, file_path in file_paths.items():
    # Validate file path
    if not os.path.exists(file_path):
        logging.error(f"File not found: {file_path}")
        raise FileNotFoundError(f"File not found: {file_path}")

    # Load all sheets, filtering only month-named sheets
    df_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")
    valid_months = [
        "January", "February", "March", "April", "May", "June", 
        "July", "August", "September", "October", "November", "December"
    ]
    df_sheets = {k: v for k, v in df_sheets.items() if k in valid_months}

    # Open the existing Excel workbook to modify only specific cells
    wb = load_workbook(file_path)

    # Process each sheet
    for sheet_name, df in df_sheets.items():
        try:
            print(f"Processing {year} - Sheet: {sheet_name} - First few date values:\n", df["Date Proposal Submitted"].head())
            
            # Check for required columns
            required_columns = ["Date Proposal Submitted", "Last Correspondence", "Contact Email", "Contact", "Project", "Value", "Won", "Lost", "Re-Bid", "Follow-Up Stage"]
            if "Follow-Up Stage" not in df.columns:
                df["Follow-Up Stage"] = 0  # Initialize if missing
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logging.error(f"Missing columns in sheet {sheet_name}: {missing_columns}")
                continue

            ws = wb[sheet_name]  # Load the worksheet
            last_correspondence_col = df.columns.get_loc("Last Correspondence") + 1  # Excel is 1-based index
            follow_up_stage_col = df.columns.get_loc("Follow-Up Stage") + 1  # Excel is 1-based index

            for index, row in df.iterrows():
                try:
                    # Validate and process row data
                    if isinstance(row["Date Proposal Submitted"], (int, float)):
                        submission_date = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=row["Date Proposal Submitted"])
                    elif isinstance(row["Date Proposal Submitted"], str):
                        submission_date = pd.to_datetime(row["Date Proposal Submitted"], errors='coerce')
                    else:
                        submission_date = pd.to_datetime(row["Date Proposal Submitted"], errors='coerce')

                    # Check if the date is valid
                    if pd.isna(submission_date) or submission_date.year < 2000:
                        logging.info(
                            f"Skipping row {index} in sheet {sheet_name}: Invalid or missing submission date ({row['Date Proposal Submitted']}) | Row data: {row.to_dict()}"
                        )
                        continue  # Skip rows with invalid submission dates
                    
                    submission_date_str = submission_date.strftime('%m-%d-%Y')

                    # Contact email extraction (robust)
                    contact_email = str(row.get("Contact Email", "") or "").strip()
                    if not contact_email:
                        logging.info(
                            f"Skipping row {index} in sheet {sheet_name}: Missing contact email | Row data: {row.to_dict()}"
                        )
                        continue  # Skip rows without contact email

                    # Stop follow-ups if project is won, lost, or set for re-bid
                    if row.get("Won", "") == "X" or row.get("Lost", "") == "X" or row.get("Re-Bid", "") == "X":
                        logging.info(
                            f"Skipping row {index} in sheet {sheet_name}: Project marked as won/lost/re-bid | Row data: {row.to_dict()}"
                        )
                        continue

                    # Determine follow-up stage
                    follow_up_stage = int(row.get("Follow-Up Stage", 0))
                    last_correspondence = pd.to_datetime(row["Last Correspondence"], errors='coerce').date() if pd.notna(row["Last Correspondence"]) else None

                    # Updated follow-up logic
                    if follow_up_stage == 0 and (not last_correspondence or (today - submission_date.date()).days >= 7):
                        days_required = 7
                        template_index = 0
                    elif follow_up_stage in [1, 2] and last_correspondence and (today - last_correspondence).days >= 14:
                        days_required = 14
                        template_index = follow_up_stage
                    elif follow_up_stage >= 3 and last_correspondence and (today - last_correspondence).days >= 14:
                        days_required = 14
                        template_index = 2  # Always use the last template
                    else:
                        # NEW: Log reason for skipping due to recent correspondence or not ready for follow-up
                        if follow_up_stage == 0:
                            reason = f"Not enough days since submission ({(today - submission_date.date()).days} < 7)"
                        elif last_correspondence:
                            reason = f"Not enough days since last correspondence ({(today - last_correspondence).days} < {days_required if 'days_required' in locals() else 'N/A'})"
                        else:
                            reason = "Unknown reason"
                        logging.info(
                            f"Skipping row {index} in sheet {sheet_name}: Not ready for follow-up. Reason: {reason} | Row data: {row.to_dict()}"
                        )
                        continue  # Skip if criteria are not met

                    # Select email template based on stage
                    email_templates = [
                        ("Quick Follow-Up on {project} Proposal", "Hi {contact},<br><br> I hope you're doing well! I wanted to follow up on our proposal for {project}, which we submitted on {date} for ${value}.<br><br> Were we competitive on pricing? Did our scope cover all the miscellaneous steel items you were expecting?<br><br> Let me know if there's anything we can clarify or adjust—we’d love to be a part of this project.<br><br>"),
                        ("Checking in on {project}", "Hi {contact},<br><br> I wanted to follow up on the status of the {project} project.<br><br> How's this project coming along? Is there anything we can do to help?<br><br>"),
                        ("Checking in again on {project}", "Hi {contact},<br><br> I wanted to check in again on the status of the {project} project.<br><br> Is this project still moving forward? Let us know if we can assist in any way.<br><br>")
                    ]

                    # Use template_index determined above
                    if template_index < 0 or template_index >= len(email_templates):
                        logging.info(f"Skipping row {index} in sheet {sheet_name}: Follow-Up Stage {follow_up_stage} out of range")
                        continue

                    subject, body = email_templates[template_index]
                    subject = subject.format(project=row['Project'])

                    try:
                        # Safely extract the first name from the Contact column
                        contact_name = row.get("Contact", "").strip()
                        first_name = contact_name.split()[0] if contact_name else "there"

                        # Value formatting (robust)
                        try:
                            value_str = f"{float(row['Value']):,.2f}"
                        except Exception:
                            value_str = str(row.get('Value', ''))

                        # Format the email body
                        body = body.format(contact=first_name, project=row['Project'], date=submission_date_str, value=value_str)

                        # Send email via Outlook MAPI
                        mail = outlook.CreateItem(0)
                        mail.To = contact_email
                        mail.Subject = subject
                        mail.HTMLBody = body + f"Looking forward to your thoughts!<br><br>{signature}"
                        mail.Send()
                        logging.info(f"Email sent to {contact_email} for project {row['Project']}")

                        # Add a 1-second delay between emails
                        time.sleep(1)

                        # Update Last Correspondence and Follow-Up Stage
                        ws.cell(row=index + 2, column=last_correspondence_col, value=today.strftime('%m-%d-%Y'))
                        # For stage 3 and above, keep stage at 3; otherwise, increment
                        if follow_up_stage >= 3:
                            ws.cell(row=index + 2, column=follow_up_stage_col, value=3)
                        else:
                            ws.cell(row=index + 2, column=follow_up_stage_col, value=follow_up_stage + 1)

                    except Exception as e:
                        logging.error(f"Error processing row {index} in sheet {sheet_name}: {str(e)}")

                except Exception as e:
                    logging.error(f"Error processing row {index} in sheet {sheet_name}: {str(e)} | Row data: {row.to_dict()}")

        except Exception as e:
            logging.error(f"Error processing sheet {sheet_name} in file {file_path}: {str(e)}")

    # Save changes to the workbook
    wb.save(file_path)  # Save changes after processing all rows in the workbook
    logging.info(f"Spreadsheet for {year} updated with last correspondence dates and follow-up stages.")