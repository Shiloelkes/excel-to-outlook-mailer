# excel-to-outlook-mailer
A professional Python-based desktop automation tool designed to streamline the process of sending personalized emails with PDF attachments via Microsoft Outlook. This script integrates Excel data to dynamically generate email content and pair records with local files.

## Key Features
- **Excel Data Management:** Automatically reads recipient details (names, emails, flags) from `.xlsx` files using Pandas.
- **Dynamic File Attachment:** Automatically locates and attaches PDF files named after specific data fields (e.g., last names).
- **Conditional Messaging:** Supports multiple email templates based on custom triggers in your spreadsheet.
- **Professional Formatting:** Uses HTML to ensure consistent styling (fonts, colors, line breaks) across different email clients.
- **Comprehensive Logging:** Generates a detailed execution report (`Email_Dispatch_Report.txt`) for tracking successes and identifying missing files.
- **User-Friendly GUI:** Simple file and folder selection dialogs powered by `tkinter`.

## Excel File Structure
To ensure the script works correctly, your Excel file (`.xlsx`) must follow this column mapping:

| Column | Data Field | Description |
| :--- | :--- | :--- |
| **C** | First Name | Used for the email greeting (e.g., "Hello John"). |
| **N** | Message Flag | Enter **'V'** to send the "Action Required" template. Leave empty for the standard report. |
| **O** | Skip Flag | Enter **'V'** to skip this row entirely during execution. |
| **R** | Last Name | Used for the greeting and to find the attachment (e.g., `Smith.pdf`). |
| **U** | Email | The recipient's email address. |

## Prerequisites
- **Operating System:** Windows (required for `pywin32` Outlook integration).
- **Email Client:** Microsoft Outlook installed and configured.
- **Python:** Version 3.7 or higher.

## Setup & Installation

1. Clone or download this repository.
2. Install the necessary Python libraries:
   ```bash
   pip install -r requirements.txt
   ```

##   Usage Instructions
Prepare your Data: Organize your Excel file according to the Excel File Structure section above.

Prepare your Attachments: Place all PDF files in a single folder. Files should be named after the recipient's last name (e.g., Smith.pdf).

## Run the Script:

```bash
python bulk_email_sender.py
```
Follow the Prompts: Select your Excel file and the folder containing your PDFs when prompted.

Review Results: Once the process finishes, check the generated report in your Excel file's folder.

## Customization
You can easily modify the script's body variables to change the email text or adjust the row.iloc indices in the code to match a different Excel layout.

Note: This tool is intended for professional automation. Always ensure you are compliant with local anti-spam regulations when sending bulk communications.
