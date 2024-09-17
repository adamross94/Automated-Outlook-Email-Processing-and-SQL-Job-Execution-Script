import win32com.client  # For Outlook automation
import os  # For file and directory operations
import pyodbc  # For connecting to SQL Server
from datetime import datetime, timedelta  # For date and time functions
import openpyxl  # For handling Excel files
import logging  # For logging script activities to a file
import time

# Define a FileHandler for logging to a file
file_handler = logging.FileHandler("script_log.log")
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(logging.Formatter(
    "%(asctime)s %(levelname)s:%(message)s"))

# Define a StreamHandler for console output
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(logging.Formatter(
    "%(asctime)s %(levelname)s:%(message)s"))

# Configure logging with both File and Console handlers
logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s %(levelname)s:%(message)s",
                    handlers=[file_handler, console_handler])  

# Initialize a list for storing summary data
summary = []

RETRY_LIMIT = 3  # Maximum number of attempts
RETRY_DELAY = 5  # Delay in seconds between retries

# Utility function to ensure a directory exists
def ensure_directory_exists(path):
    directory = os.path.dirname(path)
    if not os.path.exists(directory):
        os.makedirs(directory)
        logging.info(f"Created directory: {directory}")

# Utility function to check if the file path is valid and accessible
def is_valid_path(path):
    MAX_PATH_LENGTH = 260
    return len(path) < MAX_PATH_LENGTH and os.access(os.path.dirname(path), os.W_OK)

# Define target sender names (replace these with actual target sender names)
target_sender_names = [
    "SENDER_1",  # Example: "DOE, John (COMPANY NAME)"
    "SENDER_2",
    "SENDER_3"
    # Add more sender names here
]

# Function to remove formulas from an Excel file and save it as values only
def remove_formulas_and_save(attachment_path):
    logging.info(f"Opening workbook {attachment_path} to replace formulas with values.")
    wb = openpyxl.load_workbook(attachment_path, data_only=True)
    wb.save(attachment_path)  # Save it back to overwrite formulas
    logging.info(f"Saved workbook {attachment_path} with values only.")

# Function to process emails and attachments from Outlook
def process_emails_and_attachments(outlook, target_sender_names):
    logging.info("Entering process_emails_and_attachments function")
    try:
        # Access the inbox (replace with actual folder names)
        inbox = outlook.Folders["YOUR_OUTLOOK_ACCOUNT"].Folders["Inbox"]
        logging.info("Inbox accessed")

        # Set start and end date for filtering emails
        start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = datetime.now()

        # Define a restriction to fetch emails received within the date range
        restriction = "[ReceivedTime] >= '" + start_date.strftime("%d/%m/%Y %H:%M %p") + \
                      "' AND [ReceivedTime] <= '" + end_date.strftime("%d/%m/%Y %H:%M %p") + "'"
        logging.info(f"Restriction set for emails from {start_date} to {end_date}")

        # Retrieve filtered emails
        filtered_emails = inbox.Items.Restrict(restriction)
        filtered_emails.Sort("[ReceivedTime]", True)  # Sort by received time, descending
        logging.info(f"Number of emails fetched: {len(filtered_emails)}")

        # Loop through emails and process attachments
        for message in filtered_emails:
            sender_name = message.Sender.Name if message.Sender else ""
            logging.info(f"Sender Name: {sender_name}")

            if sender_name not in target_sender_names:
                logging.info(f"Email from {sender_name} is not in the target sender list.")
                continue

            for attachment in message.Attachments:
                logging.info(f"Checking attachment: {attachment.FileName}")
                if "TARGET_STRING" in attachment.FileName:  # Replace with actual condition
                    date_str = attachment.FileName.split(" ")[-1].replace(".xlsx", "")
                    formatted_date_str = f"{date_str[4:]}{date_str[2:4]}{date_str[:2]}"
                    attachment_path = f"C:\\path\\to\\save\\{formatted_date_str}.xlsx"  # Update this path

                    ensure_directory_exists(attachment_path)
                    if not is_valid_path(attachment_path):
                        logging.error(f"Invalid path: {attachment_path}")
                        continue

                    for attempt in range(1, RETRY_LIMIT + 1):
                        try:
                            # Save the attachment
                            attachment.SaveAsFile(attachment_path)
                            remove_formulas_and_save(attachment_path)

                            # Load the saved workbook to further process
                            wb = openpyxl.load_workbook(attachment_path, data_only=True)
                            ws = wb.active
                            ws.delete_rows(1, 1)  # Example modification

                            # Further customization here (e.g., formatting, freezing panes)
                            ws.freeze_panes = "A2"
                            wb.save(attachment_path)
                            break  # Exit retry loop on success

                        except Exception as save_error:
                            logging.error(f"Attempt {attempt} failed: {save_error}")
                            time.sleep(RETRY_DELAY)

                    if attempt == RETRY_LIMIT:
                        logging.error(f"Failed to save after {RETRY_LIMIT} attempts.")

    except Exception as e:
        logging.error(f"Error in processing: {e}")

# Function to execute a SQL Server job (update connection details)
def execute_sql_job(job_name):
    conn = None
    try:
        conn = pyodbc.connect(
            "DRIVER={ODBC Driver 17 for SQL Server};SERVER=SERVER_NAME;DATABASE=DATABASE_NAME;Trusted_Connection=yes;")
        cursor = conn.cursor()

        cursor.execute(f"EXEC msdb.dbo.sp_start_job @job_name = N'{job_name}'")
        logging.info(f"SQL job '{job_name}' started. Verify job status in SQL Server Management Studio.")

    except pyodbc.Error as e:
        logging.error(f"SQL error: {e}")

    finally:
        if conn:
            conn.close()
            logging.info("SQL connection closed.")

# Main program execution
try:
    start_time = datetime.now()
    summary.append(f"Script started at {start_time}")

    # Initialize Outlook Application and Namespace
    outlook_application = win32com.client.Dispatch("Outlook.Application")
    outlook_namespace = outlook_application.GetNamespace("MAPI")

    # Execute core functions
    process_emails_and_attachments(outlook_namespace, target_sender_names)
    # execute_sql_job('YOUR_SQL_JOB_NAME') // Uncomment and rename to trigger your SQL job. 

except pyodbc.Error as e:
    logging.error(f"Database error: {e}")
    summary.append(f"Database error: {e}")
except Exception as e:
    logging.error(f"An error occurred: {e}")
    summary.append(f"Error: {e}")
