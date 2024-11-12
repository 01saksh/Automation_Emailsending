import pandas as pd
import win32com.client as win32
import win32api

def automate_email_sending(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, engine='openpyxl')

        # Ensure the 'Status' column is of type string, filling NaN with empty strings first
        df['Status'] = df['Status'].fillna('').astype(str)

        # Initialize Outlook
        outlook = win32.Dispatch('outlook.application')
        emails_sent = 0

        for index, row in df.iterrows():
            to_email = str(row['To']).strip()  # Ensure it's a string
            subject = str(row['Subject']).strip()
            body = str(row['Body']).strip()
            client_name = str(row['Client Name']).strip()

            if df.at[index, 'Status'].lower() == 'sent':
                continue  # Skip if already sent

            # Validate email addresses
            if not to_email or "@" not in to_email:
                print(f"Invalid email for {client_name}. Skipping.")
                continue

            # Create an email
            mail = outlook.CreateItem(0)
            mail.To = to_email
            mail.Subject = subject
            mail.Body = body

            # Optional: Add CC
            if pd.notna(row['CC']):
                cc_email = str(row['CC']).strip()
                if "@" in cc_email:  # Ensure valid email address
                    mail.CC = cc_email

            # Send the email
            try:
                mail.Send()
                print(f'Email sent to {client_name} at {to_email}')
                df.at[index, 'Status'] = 'Sent'  # Update the status to 'Sent'
                emails_sent += 1
            except Exception as email_error:
                print(f"Failed to send email to {client_name}: {email_error}")
                df.at[index, 'Status'] = 'Failed'  # Mark as failed if an error occurs

        # Save the updated DataFrame back to Excel
        df.to_excel(file_path, index=False)

        # Show a message box after sending all emails
        if emails_sent > 0:
            win32api.MessageBox(0, f"Successfully sent {emails_sent} emails.", "Email Automation", 0x00001000)
        else:
            win32api.MessageBox(0, "No emails were sent.", "Email Automation", 0x00001000)

    except Exception as e:
        win32api.MessageBox(0, f"An error occurred: {e}", "Email Automation", 0x00001000)


# Example usage: Run the script immediately
excel_file_path = "C:\\Users\\ANAROCK\\AppData\\Local\\Programs\\Python\\Python312\\Automation_Codes\\Python_userconfirmationtesting.xlsx"
automate_email_sending(excel_file_path)

