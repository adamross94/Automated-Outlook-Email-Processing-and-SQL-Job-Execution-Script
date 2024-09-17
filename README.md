# Automated Outlook Email Processing and SQL Job Execution Script

This Python script automates the process of fetching emails and attachments from an Outlook inbox, processes Excel attachments by removing formulas and applying formatting, and executes a SQL Server job. It includes logging functionalities for monitoring and troubleshooting, with both file and console output.

## Key Features
- Outlook Email Automation: Connects to an Outlook inbox, filters emails based on the sender and date range, and processes specific attachments.
- Excel File Processing: Automatically saves Excel attachments, removes formulas, applies formatting, and adjusts column widths, row heights, and data types.
- SQL Job Execution: Runs a SQL Server job by executing a stored procedure through a trusted connection.
- Logging: Provides detailed logging with both file and console handlers for easy tracking of script activity and errors.
- Customizable Parameters: Easily configurable email sender list, file paths, and SQL job name to adapt the script to different use cases.

## Customization:

Users can modify:
- The Outlook inbox and folder names to target specific email accounts.
- The target sender names list to focus on specific individuals or departments.
- The attachment filename filter to process specific file types.
- The SQL job name to run their own jobs.
- File paths for saving attachments and logging.
