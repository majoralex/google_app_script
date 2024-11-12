# Google Sheets Email Automation

A simple, free email automation solution using Google Sheets and Apps Script. Send personalized emails to multiple recipients, track delivery status, and support multiple languages—all from within Google Sheets.

https://docs.google.com/spreadsheets/d/1PuIJ-JGZPsnQ5J0FWaUTClge4voLJlMtjAkhlAS4y74/edit?gid=610081433#gid=610081433


## Features
- Personalized email sending using spreadsheet data
- HTML email template support with customizable styling
- Multi-language support
- Delivery tracking and error logging
- Custom menu integration
- No external dependencies

## Setup
1. Create a new Google Sheet with two sheets:
   - "EmailList" for recipient data
   - "Log" for tracking email status

2. Add Apps Script files:
   - Navigate to Extensions > Apps Script
   - Create main script file with email sending logic
   - Add `email.html` template file

## Required Columns (EmailList sheet)
- Column A: Email addresses
- Column B: First names
- Column C: Last names
- Column D: Company names
- Column E: Language preference

## Usage
1. Fill in recipient data in the EmailList sheet
2. Click "Email Automation" in the custom menu
3. Select "Send Emails"
4. Authorize the script when prompted
5. Confirm sending when prompted

## Logging
The Log sheet automatically tracks:
- Recipient email address
- Sending status
- Timestamp
- Error messages (if any)

## Authorization
First-time usage requires authorization for:
- Google Sheets access
- Gmail sending permissions

## Contributing
Feel free to submit issues and enhancement requests.

## License
This project is released under the MIT License.

## Disclaimer
This tool uses your Gmail quota for sending emails. Please review Google's sending limits and terms of service before use.
