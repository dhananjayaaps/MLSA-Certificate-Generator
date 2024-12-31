import pandas as pd
from docx import Document
from lxml import etree
import os
from win32com.client import Dispatch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Function to generate certificates
def generate_certificate(template_path, output_path, name):
    document = Document(template_path)
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    for content_control in document.element.findall('.//w:sdt', namespaces):
        alias = content_control.find('.//w:alias', namespaces)
        if alias is not None:
            title = alias.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            text_element = content_control.find('.//w:t', namespaces)
            if title == "Name" and text_element is not None:
                text_element.text = name
            # elif title == "LinkControl" and text_element is not None:
            #     text_element.text = verification_link

    document.save(output_path)
    print(f"Certificate generated for {name}: {output_path}")

# Function to convert Word to PDF
def convert_to_pdf(docx_path, pdf_path):
    word = Dispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(docx_path))
    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    doc.Close()
    word.Quit()
    print(f"Converted to PDF: {pdf_path}")

# Function to send email with attachment
def send_email_with_attachment(from_name, from_email, to_email, subject, body, attachment_path, smtp_server, smtp_port, username, password):
    msg = MIMEMultipart()
    msg['From'] = from_name
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
        msg.attach(part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(username, password)
            server.sendmail(from_email, to_email, msg.as_string())
            print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {str(e)}")

# Main function to process Excel and send certificates
def process_and_send_certificates(excel_path, template_path, output_folder, smtp_details):
    data = pd.read_excel(excel_path)
    for index, row in data.iterrows():
        name = row['Name']
        email = row['Email']
        # verification_link = f"https://example.com/verify/{index+1}"
        
        docx_path = os.path.join(output_folder, f"{name}_certificate.docx")
        pdf_path = os.path.join(output_folder, f"{name}_certificate.pdf")
        
        generate_certificate(template_path, docx_path, name)
        convert_to_pdf(docx_path, pdf_path)
        
        send_email_with_attachment(
            from_name=smtp_details['name'],
            from_email=smtp_details['email'],
            to_email=email,

            subject=smtp_details['subject'],
            body=f"Dear {name},\n\n{smtp_details['body']}",
            
            attachment_path=pdf_path,
            smtp_server=smtp_details['server'],
            smtp_port=smtp_details['port'],
            username=smtp_details['username'],
            password=smtp_details['password']
        )

# Example Usage
excel_path = "recipients.xlsx"  # Path to the Excel file
template_path = "certificate.docx"  # Path to the Word template
output_folder = "certificates"  # Folder to save generated certificates
os.makedirs(output_folder, exist_ok=True)

smtp_details = {
    'name': "Your Name",
    'email': "Your Email",
    'server': "smtp.office365.com",
    'port': 587,
    'username': "Your Email",
    'password': "Your Password",
    'subject': "Your Subject",
    'body': "Please find your certificate attached.\n\nBest Regards,\nYour MLSA Team"
}

process_and_send_certificates(excel_path, template_path, output_folder, smtp_details)
