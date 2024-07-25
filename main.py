import pandas as pd
import os
import tempfile
import subprocess
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Function to create an email with attachment and open it in the default mail application
def open_email_app(receiver_email, subject, body, attachment_path):
    sender_email = "coolparthg11@gmail.com"  # Replace with your email

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    # Attach the PDF file
    filename = os.path.basename(attachment_path)
    with open(attachment_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {filename}',
    )

    msg.attach(part)

    # Save the email to a temporary file
    temp_dir = tempfile.gettempdir()
    eml_file_path = os.path.join(temp_dir, "email.eml")
    with open(eml_file_path, 'w') as f:
        f.write(msg.as_string())

    # Open the .eml file with the default mail application
    if os.name == 'posix':
        subprocess.call(['open', eml_file_path])  # macOS
    elif os.name == 'nt':
        os.startfile(eml_file_path)  # Windows
    elif os.name == 'os2':
        subprocess.call(['xdg-open', eml_file_path])  # Linux

# Function to read CSV and compare with PDF file name
def process_csv_and_send_emails(csv_file, pdf_directory):
    # Read CSV file
    df = pd.read_csv(csv_file, header=None, names=['Name', 'Email'])

    # Iterate over the CSV rows
    for index, row in df.iterrows():
        name = row['Name']
        email = row['Email']
        
        # Generate the expected PDF file name
        pdf_file_name = f"{name.replace(' ', '')}.pdf"
        pdf_file_path = os.path.join(pdf_directory, pdf_file_name)

# Example usage
csv_file = 'letters.csv'  # Path to your CSV file
pdf_directory = os.path.join(os.getcwd(), 'InternshipAppReader')  # Path to the directory containing PDF files
process_csv_and_send_emails(csv_file, pdf_directory)
