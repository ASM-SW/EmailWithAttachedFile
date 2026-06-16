# Email With Attached File Help

## Overview
This application automates sending personalized emails with unique attachments. It is designed to work with Microsoft 365 / Outlook using OAuth2 authentication. The main screen shows you important settings for sending emails.  If you would like to change any of them click on the "Change Settings" button.

## Sending Emails
Click the **Start** button to begin the process. The application will prompt for authentication via your web browser if a valid OAuth2 session is not already established.

A log in the main window will display the real-time status of each email. Once the process completes or is cancelled, a CSV results file (e.g., `input_out.csv`) is generated in the same folder as your input file. This file contains all original data along with **Status** and **Message** columns indicating the outcome for each email.
To run the program and send emails click on Start.

## Setup and Configuration
Click the **Settings** button to configure your environment. 
At the top of the settings window are the following fields:
- **Template File**: A `.txt` file used for the email body. You can use the `<NAME>` placeholder, which will be replaced by the recipient's name from the CSV.
- **Input File**: This is a CSV file.  It had a header row and one row for each email recipient.  See "CSV Column Mapping" for a list of required columns.
- **Mail Subject**: The subject line that will be used for all outgoing emails.

### Account Information
- **Tenant ID**: The Directory (tenant) ID from your Azure App Registration.
- **Client ID**: The Application (client) ID from your Azure App Registration.
- **Sender Email**: The Microsoft 365 / Outlook email address used to send the messages.
- **Sender Name**: The display name that recipients will see in the "From" field.

### CSV Column Mapping
The CSV file has a header row.  The following entries are where you put the column name from the headings rowl
- **Name Column** The Name of the person receiving the email.  It is used to replace teh <NAME> token in the email template file.
- **Email Column** A list of email addresses to send the email to.  The email addresses can be separate by comma, semicolon or space. All email addresses are added to the email To field.
-**Filename Column** This is the name of the file to send to this specific recipient.

#### ** Attachments**: A list of files that will be attached to every single email sent, in addition to the unique file specified in the CSV.

## Settings File
All of the settings on the settings page is saved into a json file.  It is located at
```
C:/Users/XXXX/AppData/Local/EmailWithAttachedFile/EmailWithAttachedFile.json 
   where XXXX is your abbreviated username.
````
