import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

# Read the emails from an Excel sheet or CSV
def read_email_list(file_path):
    df = pd.read_excel(file_path)  # Or pd.read_csv('file.csv')
    return df

# Update the status column in the Excel sheet
def update_status(file_path, recipient_email, status):
    df = pd.read_excel(file_path)

    # Add or update the Status column
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    df.loc[df['Email'] == recipient_email, 'Status'] = f"{status} at {timestamp}"

    # Save the updated DataFrame back to the Excel file
    df.to_excel(file_path, index=False)

# Function to send an email with an attachment
def send_email(smtp_server, smtp_port, sender_email, sender_password, recipient_email, file_path, subject, body, attachment_path=None):
    try:
        # Set up the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure the connection
        server.login(sender_email, sender_password)

        # Create the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Attach the body of the email
        msg.attach(MIMEText(body, 'plain'))

        # Attach the file if an attachment path is provided
        if attachment_path:
            with open(attachment_path, 'rb') as file:
                part = MIMEApplication(file.read(), Name=attachment_path.split("/")[-1])
                part['Content-Disposition'] = f'attachment; filename="{attachment_path.split("/")[-1]}"'
                msg.attach(part)

        # Send the email
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"Email sent to {recipient_email}")
        
        # Update the status in the Excel file
        update_status(file_path, recipient_email, "Sent")

        # Quit the server
        server.quit()
    except Exception as e:
        print(f"Error sending email to {recipient_email}: {e}")
        # Update the status in the Excel file for failed email
        update_status(file_path, recipient_email, f"Failed ({str(e)})")

def main():
    # Your email credentials
    sender_email = ""
    sender_password = ""
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # Email subject and body
    subject = 'Invitation to Register for Six-Day Online ATAL-Sponsored FDP on "Generative AI: Techniques, Tools, and Applications."'
    body = """Dear All,
We are delighted to announce that the Department of Computer Science and Engineering, at National Institute of Technology, Delhi, is organizing a six-day online Faculty Development Programme (FDP) sponsored by AICTE-ATAL from February 03th to 08th, 2025.

Theme of the FDP:

"Generative AI: Techniques, Tools, and Applications."

This FDP is free of cost for all participants and is designed to offer valuable insights into the role of AI and emerging technologies in building a sustainable future.

We kindly request you to share this information with faculty members, research scholars, and other interested participants.

How to Register:

 Participants can register at:

👉 https://atalacademy.aicte-india.org/login

Steps to Register:

 1. Sign up as a participant and fill in your details.

2. Go to FDPs → ATAL → January → ENGINEERING → Online.

 (Tip: Use Ctrl+F and search for FDP Application No: 1730463455.)

3. Select the Institute: National Institute of Technology, Delhi.

4. Register for the FDP titled:

"Generative AI: Techniques, Tools, and Applications"

(Application No: 1730463455).

We look forward to your participation and support!

Contact Information

For any queries, please reach out via WhatsApp or text:  
📞9896592476, 9554741467

The program brochure and schedule are attached for your reference. 

Best regards,

  
FDP Co-ordinator 

Dr.Anurag Singh- 7835014014

Associate Professor - Department of Computer Science and Engineering

 

FDP Co-Coordinator

Dr. Karan Verma- 8003389258

Associate Professor - Department of Computer Science and Engineering"""

    # Attachment path (update with the correct path to your attachment file)
    attachment_path = "6 DAY online FACULTY DEVELOPMENT PROGRAM  ON  “Generative AI Techniques, Tools, and Applications” (5).pdf"  # Adjust this

    # Read emails from an Excel or CSV file
    file_path = "testSheet.xlsx"  # Update with your file path
    df = read_email_list(file_path)

    # Iterate through each email and send the message
    for recipient_email in df['Email']:
        send_email(smtp_server, smtp_port, sender_email, sender_password, recipient_email, file_path, subject, body, attachment_path)

if __name__ == "__main__":
    main()
