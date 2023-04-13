
# Outlook Email Auto Sender


This is a Python script for sending emails using Microsoft Outlook. The script reads data from an Excel file and extracts text from a Word document, and then uses the extracted data to create an email in Outlook.




## Prerequisites

To use this script, you need to have the following software installed on your machine:

- Python 3.x
- Microsoft Excel
- Microsoft Word
- Microsoft Outlook

You also need to have the win32com and pandas Python packages installed. You can install these packages using pip:

```bash
pip install win32com pandas
```
## Usage/Examples

1. Open the Excel file containing the email addresses and subject lines. Make sure that the sheet containing the relevant data is named "Sheet1".

2. Make sure that the Word document containing the email body text is named "Auto.docx".

3. Close the Excel file.

4. Run the Python script using the command:
```bash
python Outlook_Automated.py
```

5. Outlook will open, with a new email message populated with the email addresses and subject lines from the Excel file, and the body text from the Word document.

6. You can then edit the email as required, and send it.

## Functionality
The script reads the email addresses and subject lines from the first sheet of the Excel file. The column containing the email addresses must be named "To_Email", and the column containing the cc email addresses must be named "CC_Email". The email body text is extracted from the Word document named "Auto.docx".

The email is created using the EzMail function, which uses the win32com package to create an Outlook email object. The email addresses and subject line are set using the To and Subject properties of the email object. The email body text is inserted using the GetInspector.WordEditor.Range(Start=0, End=0).Paste() method.

