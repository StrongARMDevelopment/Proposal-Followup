import os
import pandas as pd
import datetime
import logging
import win32com.client as win32
from openpyxl import load_workbook
import time
import configparser
import sys
import zipfile
import ast
import shutil

# --- Global Script Constants ---
CONFIG_FILE_NAME = "config.ini"
SCRIPT_VERSION = "2.0.0"
EXCEL_DATE_OFFSET = datetime.datetime(1899, 12, 30)
EXCEL_ROW_OFFSET = 2  # Excel rows are 1-based, plus header

def excel_date_to_datetime(excel_date):
    """Convert Excel serial date to datetime."""
    if isinstance(excel_date, (int, float)) and excel_date > 0:
        return EXCEL_DATE_OFFSET + datetime.timedelta(days=excel_date)
    return pd.to_datetime(excel_date, errors='coerce')

def create_lock_file(lock_file_path):
    """Create a lock file to prevent concurrent script runs."""
    global lock_file_handle
    try:
        lock_file_handle = open(lock_file_path, 'x')
        logging.info(f"Lock file created: {lock_file_path}")
        return True
    except FileExistsError:
        logging.error(f"Lock file {lock_file_path} already exists. Another instance may be running.")
        return False
    except IOError as e:
        logging.error(f"Failed to create lock file {lock_file_path}: {e}")
        return False

def remove_lock_file(lock_file_path):
    """Remove the lock file on script exit."""
    global lock_file_handle
    try:
        if lock_file_handle:
            lock_file_handle.close()
            lock_file_handle = None
        if os.path.exists(lock_file_path):
            os.remove(lock_file_path)
            logging.info(f"Lock file removed: {lock_file_path}")
    except IOError as e:
        logging.error(f"Failed to remove lock file {lock_file_path}: {e}")

def load_configuration():
    """Load and parse the configuration file."""
    parser = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE_NAME):
        logging.critical(f"CRITICAL: Configuration file '{CONFIG_FILE_NAME}' not found. Exiting.")
        raise FileNotFoundError(f"Configuration file '{CONFIG_FILE_NAME}' not found.")
    try:
        parser.read(CONFIG_FILE_NAME)
        # Convert specific string settings from config to appropriate types
        parser.set('Settings', 'DryRun', str(parser.getboolean('Settings', 'DryRun', fallback=False)))
        parser.set('Settings', 'YearsToProcess', str([year.strip() for year in parser.get('Settings', 'YearsToProcess', fallback='').split(',') if year.strip()]))
        parser.set('Settings', 'ValidMonths', str([month.strip() for month in parser.get('Settings', 'ValidMonths', fallback='').split(',') if month.strip()]))
        logging.info(f"Configuration loaded from '{CONFIG_FILE_NAME}'.")
        return parser
    except configparser.Error as e:
        logging.critical(f"CRITICAL: Error parsing configuration file '{CONFIG_FILE_NAME}': {e}. Exiting.")
        raise

def initialize_outlook(config):
    """Initialize Outlook application and retrieve signature."""
    global outlook, namespace, outlook_signature
    if config.getboolean('Settings', 'DryRun'):
        logging.info("DRY RUN: Skipping Outlook initialization.")
        return

    try:
        outlook = win32.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        logging.info("Outlook MAPI session initialized.")

        try:
            dummy_mail = outlook.CreateItem(0)
            if not dummy_mail.HTMLBody:
                dummy_mail.Display()
                time.sleep(0.5)
            outlook_signature = dummy_mail.HTMLBody
            dummy_mail.Close(0)
            if outlook_signature:
                logging.info("Successfully retrieved Outlook signature.")
            else:
                logging.warning("Outlook signature is empty or could not be retrieved.")
        except Exception as e:
            logging.error(f"Could not retrieve Outlook signature: {e}.")
            outlook_signature = ""
    except Exception as e:
        logging.error(f"Outlook is not available or not configured properly: {e}. Emails will not be sent.")
        outlook = None

def get_outlook_account(desired_email):
    """Get the Outlook account matching the desired email."""
    if not outlook or not desired_email:
        return None
    try:
        for acc in outlook.Session.Accounts:
            if acc.SmtpAddress.lower() == desired_email.lower():
                logging.info(f"Using specified Outlook account: {desired_email}")
                return acc
        logging.warning(f"Specified Outlook account '{desired_email}' not found. Using default.")
    except Exception as e:
        logging.error(f"Error trying to get specified Outlook account: {e}")
    return None

def process_proposals(config, today_date):
    """Process proposals from Excel files and group follow-ups by contact email."""
    grouped_follow_ups = {}
    cols = config['Columns']
    required_df_cols = [
        cols['DateProposalSubmitted'], cols['LastCorrespondence'], cols['ContactEmail'],
        cols['ContactName'], cols['ProjectName'], cols['Value'],
        cols['Won'], cols['Lost'], cols['ReBid'], cols['FollowUpStage']
    ]
    years_to_process = ast.literal_eval(config.get('Settings', 'YearsToProcess'))
    valid_months_list = ast.literal_eval(config.get('Settings', 'ValidMonths'))

    for year in years_to_process:
        try:
            file_path = config.get('Paths', f'Log{year}Path')
        except configparser.NoOptionError:
            logging.warning(f"No path configured for year {year} in '{CONFIG_FILE_NAME}'. Skipping.")
            continue

        if not os.path.exists(file_path):
            logging.error(f"File not found: {file_path} for year {year}. Skipping this file.")
            continue

        logging.info(f"Processing file: {file_path} for year {year}")
        try:
            df_sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl", usecols=lambda x: x in required_df_cols)
        except (FileNotFoundError, PermissionError, zipfile.BadZipFile, ValueError) as e:
            logging.error(f"Error reading Excel file {file_path}: {e}. Skipping.")
            continue
        except Exception as e:
            logging.error(f"Could not read Excel file {file_path} using pandas: {e}. Skipping this file.")
            continue

        sheet_column_indices_df = {}

        for sheet_name, df in df_sheets.items():
            if sheet_name not in valid_months_list:
                logging.debug(f"Skipping sheet '{sheet_name}' in {file_path} as it's not in ValidMonths list.")
                continue

            logging.info(f"Collecting data from {year} - Sheet: {sheet_name}")

            current_df_cols = df.columns.tolist()
            missing_in_df = [col for col in required_df_cols if col not in current_df_cols]
            if missing_in_df:
                logging.error(f"Required columns {missing_in_df} not found in DataFrame for sheet '{sheet_name}' (file: {file_path}) after selective loading. Skipping sheet.")
                continue

            if cols['FollowUpStage'] not in df.columns:
                logging.info(f"'{cols['FollowUpStage']}' column missing in sheet '{sheet_name}'. Initializing with 0.")
                df[cols['FollowUpStage']] = 0

            try:
                sheet_column_indices_df[sheet_name] = {
                    cols['LastCorrespondence']: df.columns.get_loc(cols['LastCorrespondence']),
                    cols['FollowUpStage']: df.columns.get_loc(cols['FollowUpStage'])
                }
            except KeyError as ke:
                logging.error(f"A required column for indexing not found in DataFrame for sheet {sheet_name}: {ke}. Skipping sheet.")
                continue

            for index, row in df.iterrows():
                try:
                    submission_date_val = row[cols['DateProposalSubmitted']]
                    submission_date = excel_date_to_datetime(submission_date_val)
                    if pd.isna(submission_date) or submission_date.year < 2000:
                        logging.debug(f"Skipping Excel row {index+EXCEL_ROW_OFFSET} in sheet {sheet_name}: Invalid submission date ({submission_date_val}).")
                        continue
                    submission_date_dt_date = submission_date.date()
                    submission_date_str = submission_date.strftime('%m-%d-%Y')

                    contact_email = str(row.get(cols['ContactEmail'], "") or "").strip().lower()
                    if not contact_email or "@" not in contact_email or "." not in contact_email:
                        logging.debug(f"Skipping Excel row {index+EXCEL_ROW_OFFSET} in sheet {sheet_name}: Missing or invalid contact email ('{row.get(cols['ContactEmail'], '')}').")
                        continue

                    if str(row.get(cols['Won'], "")).strip().upper() == "X" or \
                       str(row.get(cols['Lost'], "")).strip().upper() == "X" or \
                       str(row.get(cols['ReBid'], "")).strip().upper() == "X":
                        continue

                    follow_up_stage = int(row.get(cols['FollowUpStage'], 0))
                    last_correspondence_val = row[cols['LastCorrespondence']]
                    last_correspondence_date = None
                    if pd.notna(last_correspondence_val):
                        last_correspondence_date = excel_date_to_datetime(last_correspondence_val)
                        if pd.isna(last_correspondence_date):
                            last_correspondence_date = None
                        else:
                            last_correspondence_date = last_correspondence_date.date()

                    template_index = -1
                    ready_for_follow_up = False
                    days_first_followup = config.getint('Settings', 'DaysFirstFollowUp')
                    days_subsequent_followups = config.getint('Settings', 'DaysSubsequentFollowUps')

                    if follow_up_stage == 0:
                        relevant_date_for_calc = submission_date_dt_date
                        if last_correspondence_date:
                            relevant_date_for_calc = last_correspondence_date
                        days_since_relevant = (today_date - relevant_date_for_calc).days
                        if days_since_relevant >= days_first_followup:
                            template_index = 0
                            ready_for_follow_up = True
                    elif follow_up_stage > 0:
                        if last_correspondence_date:
                            days_since_relevant = (today_date - last_correspondence_date).days
                            if days_since_relevant >= days_subsequent_followups:
                                template_index = 1 if follow_up_stage == 1 else 2
                                ready_for_follow_up = True
                        else:
                            logging.warning(f"Skipping Excel row {index+EXCEL_ROW_OFFSET} (Project: {row.get(cols['ProjectName'], 'N/A')}): Stage is {follow_up_stage} but no valid 'Last Correspondence' date.")
                            continue

                    if not ready_for_follow_up:
                        continue

                    contact_name_full = str(row.get(cols['ContactName'], "")).strip()
                    first_name = contact_name_full.split()[0] if contact_name_full else "there"

                    value_str_formatted = ""
                    try:
                        val = row[cols['Value']]
                        if pd.notna(val) and str(val).strip() != "":
                            value_str_formatted = f"{float(val):,.2f}"
                    except (ValueError, TypeError):
                        value_str_formatted = str(row.get(cols['Value'], '')).strip()

                    proposal_details = {
                        "project_name": str(row.get(cols['ProjectName'], 'N/A')),
                        "submission_date_str": submission_date_str,
                        "value_str": value_str_formatted,
                        "contact_first_name": first_name,
                        "row_index_df": index,
                        "sheet_name": sheet_name,
                        "workbook_path": file_path,
                        "new_follow_up_stage": follow_up_stage + 1,
                        "email_snippet_key": template_index,
                        "last_correspondence_col_df_idx": sheet_column_indices_df[sheet_name][cols['LastCorrespondence']],
                        "follow_up_stage_col_df_idx": sheet_column_indices_df[sheet_name][cols['FollowUpStage']]
                    }
                    grouped_follow_ups.setdefault(contact_email, []).append(proposal_details)
                    logging.info(f"Queued follow-up for '{proposal_details['project_name']}' to {contact_email} (Sheet: {sheet_name}, Excel Row: {index+EXCEL_ROW_OFFSET}).")

                except Exception as e:
                    logging.error(f"Error processing Excel row {index+EXCEL_ROW_OFFSET} in sheet {sheet_name} ({file_path}): {e}. Row data: {row.to_dict()}", exc_info=True)

    logging.info(f"--- Phase 1 Complete: Collected {sum(len(v) for v in grouped_follow_ups.values())} total qualifying follow-ups for {len(grouped_follow_ups)} unique contacts. ---")
    return grouped_follow_ups

def send_consolidated_emails(config, grouped_follow_ups):
    """Send consolidated emails for grouped follow-ups."""
    global outlook_signature
    successful_sends_for_excel_update = []
    is_dry_run = config.getboolean('Settings', 'DryRun')

    if not outlook and not is_dry_run:
        logging.error("Outlook not initialized. Skipping Email Sending Phase.")
        return successful_sends_for_excel_update
    if not grouped_follow_ups:
        logging.info("No qualifying follow-ups collected. Skipping Email Sending Phase.")
        return successful_sends_for_excel_update

    logging.info(f"--- Phase 2: Sending Consolidated Emails {'(DRY RUN)' if is_dry_run else ''} ---")

    desired_account_email = config.get('Settings', 'DesiredOutlookAccount', fallback='').strip()
    specific_outlook_account = get_outlook_account(desired_account_email) if desired_account_email else None

    email_cfg = config['EmailBody']
    subject_cfg = config['EmailSubjects']
    snippet_cfg = config['EmailSnippets']

    for contact_email, proposals_list in grouped_follow_ups.items():
        contact_first_name = proposals_list[0]['contact_first_name']
        project_names = [p['project_name'] for p in proposals_list]

        if len(project_names) == 1:
            subject_line = subject_cfg.get('SingleProject', "Follow-up on {project_name}").format(project_name=project_names[0])
        elif len(project_names) == 2:
            subject_line = subject_cfg.get('TwoProjects', "Follow-up on {project_name_1} and {project_name_2}").format(project_name_1=project_names[0], project_name_2=project_names[1])
        else:
            subject_line = subject_cfg.get('MultipleProjects', "Follow-up on {first_project_name} & others").format(first_project_name=project_names[0])

        body_html = email_cfg.get('Greeting', "<p>Hi {contact_first_name},</p>").format(contact_first_name=contact_first_name)
        body_html += email_cfg.get('Intro', "<p>Follow-up on proposals:</p>")
        body_html += email_cfg.get('ProjectListStart', "<ul>")

        for proposal in proposals_list:
            snippet_text = snippet_cfg.get(f"Snippet_{proposal['email_snippet_key']}", "Status update requested.")
            value_info = email_cfg.get('ValueFormat', ", value: ${value_str}").format(value_str=proposal['value_str']) if proposal['value_str'] else ""

            body_html += email_cfg.get('ProjectListItem', "<li><strong>{project_name}</strong> ({submission_date_str}{value_info})<br/><em>{snippet_text}</em></li>").format(
                project_name=proposal['project_name'],
                submission_date_str=proposal['submission_date_str'],
                value_info=value_info,
                snippet_text=snippet_text
            )
        body_html += email_cfg.get('ProjectListEnd', "</ul>")
        body_html += email_cfg.get('Closing', "<p>Let us know if we can help.</p>")

        if outlook_signature:
            body_html += f"<br><br>{outlook_signature}"
        else:
            body_html += email_cfg.get('DefaultSignatureFallback', "<br><br>Best regards,")

        if is_dry_run:
            logging.info(f"DRY RUN: Would send email to {contact_email} - Subject: '{subject_line}'")
            logging.debug(f"DRY RUN: Email body for {contact_email}:\n{body_html}")
            successful_sends_for_excel_update.extend(proposals_list)
            continue

        if not outlook:
            logging.error(f"Cannot send email to {contact_email} as Outlook is not available.")
            continue

        max_attempts = config.getint('Settings', 'MaxEmailSendAttempts', fallback=3)
        retry_delay = config.getint('Settings', 'EmailRetryDelaySeconds', fallback=5)
        sent_successfully = False
        for attempt in range(1, max_attempts + 1):
            try:
                mail = outlook.CreateItem(0)
                if specific_outlook_account:
                    mail.SendUsingAccount = specific_outlook_account

                mail.To = contact_email
                mail.Subject = subject_line
                mail.HTMLBody = body_html
                mail.Send()
                logging.info(f"Consolidated email sent to {contact_email} for {len(proposals_list)} project(s) (Attempt {attempt}/{max_attempts}).")
                sent_successfully = True
                break
            except Exception as e:
                logging.warning(f"Attempt {attempt}/{max_attempts} failed to send email to {contact_email}: {e}")
                if attempt < max_attempts:
                    time.sleep(retry_delay)
                else:
                    logging.error(f"Failed to send email to {contact_email} after {max_attempts} attempts.")

        if sent_successfully:
            successful_sends_for_excel_update.extend(proposals_list)
            time.sleep(config.getint('Settings', 'EmailSendDelaySeconds', fallback=1))

    logging.info(f"--- Phase 2 Complete: Email Sending Attempted {'(DRY RUN)' if is_dry_run else ''}. ---")
    return successful_sends_for_excel_update

def update_excel_sheets(config, successful_updates, today_date):
    """Update Excel sheets for proposals that were successfully emailed."""
    is_dry_run = config.getboolean('Settings', 'DryRun')
    backup_excel = config.getboolean('Settings', 'BackupExcelBeforeSave', fallback=True)
    if not successful_updates:
        logging.info("No proposals to update in Excel. Skipping Excel Update Phase.")
        return

    logging.info(f"--- Phase 3: Updating Excel Sheets {'(DRY RUN)' if is_dry_run else ''} ---")
    updates_by_workbook = {}
    for proposal_info in successful_updates:
        updates_by_workbook.setdefault(proposal_info['workbook_path'], []).append(proposal_info)

    for wb_path, proposals_in_wb in updates_by_workbook.items():
        if is_dry_run:
            for update_item in proposals_in_wb:
                logging.info(f"DRY RUN: Would update Excel: File='{wb_path}', Sheet='{update_item['sheet_name']}', "
                             f"Excel Row={update_item['row_index_df']+EXCEL_ROW_OFFSET}, Project='{update_item['project_name']}', "
                             f"New Stage={update_item['new_follow_up_stage']}")
            continue

        try:
            logging.info(f"Attempting to update Excel file: {wb_path}")
            if backup_excel and os.path.exists(wb_path):
                backup_path = wb_path + ".bak"
                shutil.copy2(wb_path, backup_path)
                logging.info(f"Backup created: {backup_path}")

            wb = load_workbook(wb_path)
            for sheet_name, sheet_updates_list in pd.DataFrame(proposals_in_wb).groupby('sheet_name'):
                if sheet_name not in wb.sheetnames:
                    logging.error(f"Sheet '{sheet_name}' not found in workbook '{wb_path}'. Skipping updates for this sheet.")
                    continue
                ws = wb[sheet_name]
                logging.info(f"Updating sheet: '{sheet_name}' in '{wb_path}'.")

                for update_item in sheet_updates_list.to_dict('records'):
                    excel_row_num = update_item['row_index_df'] + EXCEL_ROW_OFFSET
                    last_corr_col_excel_idx = update_item['last_correspondence_col_df_idx'] + 1
                    stage_col_excel_idx = update_item['follow_up_stage_col_df_idx'] + 1
                    try:
                        ws.cell(row=excel_row_num, column=last_corr_col_excel_idx, value=today_date.strftime('%m-%d-%Y'))
                        ws.cell(row=excel_row_num, column=stage_col_excel_idx, value=update_item['new_follow_up_stage'])
                        logging.info(f"  Updated sheet '{sheet_name}', Excel row {excel_row_num} (Project: '{update_item['project_name']}').")
                    except Exception as cell_e:
                        logging.error(f"  Error updating cell in sheet '{sheet_name}', Excel row {excel_row_num}: {cell_e}", exc_info=True)

            wb.save(wb_path)
            logging.info(f"Successfully saved changes to {wb_path}")

        except FileNotFoundError:
            logging.error(f"File not found for update: {wb_path}.")
        except PermissionError:
            logging.error(f"Permission denied for updating file: {wb_path}.")
        except zipfile.BadZipFile:
            logging.error(f"File is corrupted or not a valid Excel file (openpyxl): {wb_path}.")
        except Exception as e:
            logging.error(f"Error processing or saving Excel file {wb_path} for updates: {e}", exc_info=True)
    logging.info(f"--- Phase 3 Complete: Excel Update Attempted {'(DRY RUN)' if is_dry_run else ''}. ---")

def setup_logging(config):
    """Set up logging based on config."""
    log_file_path = config.get('Paths', 'ScriptLogFile', fallback='followup_automation.log')
    log_level = getattr(logging, config.get('Settings', 'LogLevel', fallback='INFO').upper(), logging.INFO)
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    file_handler = logging.FileHandler(log_file_path)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    root_logger.addHandler(file_handler)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    root_logger.addHandler(console_handler)
    root_logger.setLevel(log_level)

def main():
    """Main script execution."""
    global outlook, namespace, outlook_signature, lock_file_handle
    today_date = datetime.datetime.today().date()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    try:
        config = load_configuration()
    except Exception as e:
        sys.exit(1)

    setup_logging(config)
    logging.info(f"--- Script Started (Version: {SCRIPT_VERSION}) ---")
    is_dry_run = config.getboolean('Settings', 'DryRun')
    if is_dry_run:
        logging.warning("IMPORTANT: Script is running in DRY RUN mode. No emails will be sent, and no Excel files will be saved.")

    lock_file_path = config.get('Paths', 'LockFilePath', fallback='followup_automation.lock')
    if not create_lock_file(lock_file_path):
        logging.critical("Exiting due to lock file issue.")
        sys.exit(1)

    try:
        initialize_outlook(config)
        collected_follow_ups = process_proposals(config, today_date)
        successfully_updated_proposals = send_consolidated_emails(config, collected_follow_ups)
        update_excel_sheets(config, successfully_updated_proposals, today_date)
    except Exception as e:
        logging.critical(f"An unhandled error occurred in main execution: {e}", exc_info=True)
    finally:
        remove_lock_file(lock_file_path)
        logging.info("--- Script Finished ---")

if __name__ == "__main__":
    main()