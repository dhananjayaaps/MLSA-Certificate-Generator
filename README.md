# Certificate Generation and Email Automation

This project automates the generation of certificates from a Word template, converts them to PDF, and emails them to recipients listed in an Excel sheet.

## Features
- Reads recipient details (names and emails) from an Excel file.
- Generates personalized certificates using a Word template.
- Converts the certificates to PDF format.
- Sends the certificates as email attachments to recipients.

## Prerequisites
- Microsoft Word (required for `pywin32` to work).
- Python 3.7 or higher.

## Installation
1. Clone the repository:
    ```bash
    git clone https://github.com/dhananjayaaps/MLSA-Certificate-Generator
    cd MLSA-Certificate-Generator
    ```
2. Install the required Python packages:
    ```bash
    pip install -r requirements.txt
    ```

## Usage
1. Prepare an Excel file (`recipients.xlsx`) with the following columns:
   - `Name`: The recipient's name.
   - `Email`: The recipient's email address.

2. Customize the Word template (`certificate.docx`) with placeholders for `Name`.

3. Update the `smtp_details` dictionary in the script with your SMTP server details and credentials.

4. Run the script:
    ```bash
    python script_name.py
    ```

5. Certificates will be generated and saved in the `certificates` folder, and emails will be sent to the recipients.

## SMTP Server Configuration
For Microsoft Outlook (Office365):
- SMTP Server: `smtp.office365.com`
- Port: `587`

Ensure you use an app-specific password or enable "less secure apps" access if required.

## Dependencies
- `pandas`: For reading the Excel file.
- `python-docx`: For manipulating Word documents.
- `lxml`: For working with XML in Word documents.
- `pywin32`: For Word to PDF conversion.
- `openpyxl`: For Excel file compatibility.

## Notes
- Ensure Microsoft Word is installed on the system where this script is run, as `pywin32` relies on it.
- Do not hard-code sensitive information (e.g., email passwords). Use environment variables or a secure vault instead.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
