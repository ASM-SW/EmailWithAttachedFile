# Email With Attached File Help

## Overview
This application automates sending personalized emails with unique attachments. It is designed to work with Microsoft 365 / Outlook using OAuth2 authentication.

## Setup and Configuration
1. **Settings**: Click the 'Settings' button to configure your environment.
   - **Tenant ID**: The Directory (tenant) ID from your Azure App Registration.
   - **Client ID**: The Application (client) ID from your Azure App Registration.
   - **Sender Email**: The Microsoft 365 / Outlook email address used to send the messages.
   - **Sender Name**: The display name that recipients will see in the "From" field.
   - **Template File**: A `.txt` file used for the email body. You can use the `<NAME>` placeholder, which will be replaced by the recipient's name from the CSV.
   - **Input CSV**: The path to the CSV file containing your list of recipients and their specific attachments.
   - **Subject**: The subject line that will be used for all outgoing emails.
   - **Common Attachments**: A list of files that will be attached to every single email sent, in addition to the unique file specified in the CSV.

## CSV Input Format
The CSV file must have a header row with the following column names:
- `NameLastFirst`: Recipient's display name.
- `Email`: Recipient's email address (supports multiple semicolon-separated values).
- `FileName`: Full path to the specific file to be attached for this recipient.

## Sending Emails
Click the **Start** button to begin. A log will display the status of each email. Once finished, a CSV results file (`..._out.csv`) will be generated in the same folder as your input file.