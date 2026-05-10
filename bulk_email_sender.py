import pandas as pd
import win32com.client as win32
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def send_bulk_emails():
    """
    A generic tool to send personalized emails with PDF attachments 
    using data from an Excel spreadsheet and Microsoft Outlook.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 1. Select the Excel file
    excel_path = filedialog.askopenfilename(title="Select the Excel file", filetypes=[("Excel", "*.xlsx")])
    if not excel_path: return

    # 2. Select the folder containing the PDF files
    pdf_folder = filedialog.askdirectory(title="Select the folder containing the PDF files")
    if not pdf_folder: return

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Cannot read the Excel file: {e}")
        return

    try:
        outlook = win32.Dispatch('outlook.application')
    except:
        messagebox.showerror("Error", "Outlook must be open in the background")
        return

    report_lines = [f"Email Dispatch Report - {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n", "-"*50 + "\n"]
    
    print("Starting to send emails... Please do not close this window.")

    for index, row in df.iterrows():
        try:
            # --- Define columns (A=0, C=2, N=13, O=14, R=17, U=20) ---
            # These can be customized based on your Excel structure
            first_name = str(row.iloc[2]).strip()    # Column C
            last_name = str(row.iloc[17]).strip()    # Column R
            email = str(row.iloc[20]).strip()        # Column U
            skip_val = str(row.iloc[14]).strip()     # Column O
            alt_msg_val = str(row.iloc[13]).strip()  # Column N

            # --- Skip condition (Column O) ---
            if skip_val.upper() == 'V':
                print(f"Skipping row {index+2} ('V' in skip column)")
                continue

            # --- Prepare file path ---
            # Matches PDF files named after the 'last_name' variable
            file_name = last_name if last_name.lower().endswith('.pdf') else f"{last_name}.pdf"
            full_file_path = os.path.join(pdf_folder, file_name)

            # Validate data and check if attachment exists
            if "@" not in email or email.lower() == 'nan' or not os.path.exists(full_file_path):
                if not os.path.exists(full_file_path) and first_name.lower() != 'nan':
                    report_lines.append(f"Missing file: {file_name} for {first_name} {last_name}\n")
                continue

            # --- Define generic subject ---
            subject = f"Important Account Update - {last_name}"

            # --- Select generic email body based on Column N ---
            if alt_msg_val.upper() == 'V':
                # Alternative message (Action Required)
                body = (
                    f"Hello {first_name},\n\n"
                    f"Please review the attached document. Action is required on your account.\n"
                    f"Kindly reply to this email with the updated forms at your earliest convenience.\n\n"
                    f"Thank you,\n"
                    f"[YOUR_COMPANY_NAME]"
                )
            else:
                # Default message (Status Report)
                body = (
                    f"Hello {first_name},\n\n"
                    f"Attached is your latest status report.\n"
                    f"Please review the document for your records. No further action is required at this time.\n\n"
                    f"If you have any questions, feel free to reach out.\n\n"
                    f"Best regards,\n"
                    f"[YOUR_COMPANY_NAME]"
                )

            # --- Create and send email ---
            mail = outlook.CreateItem(0)
            mail.To = email
            mail.Subject = subject
            
            # Convert line breaks to HTML tags for standard email formatting
            html_text = body.replace('\n', '<br>')
            mail.HTMLBody = f"""
            <div style="font-family: Arial, sans-serif; font-size: 11pt; color: #333333;">
                {html_text}
            </div>
            """
            
            mail.Attachments.Add(full_file_path)
            mail.Send()

            print(f"Email sent successfully to {first_name} {last_name}")
            report_lines.append(f"Sent: {first_name} {last_name} | {email}\n")

        except Exception as e:
            err_log = f"Error in row {index+2}: {e}"
            print(err_log)
            report_lines.append(err_log + "\n")

    # Save final report
    report_path = os.path.join(os.path.dirname(excel_path), "Email_Dispatch_Report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.writelines(report_lines)
    
    messagebox.showinfo("Finished", f"Process completed!\nReport saved at:\n{report_path}")

if __name__ == "__main__":
    send_bulk_emails()
