
# Email Automation Script

This Python script automates the process of sending personalized email invitations for the **USF ACM AI Convention 2025** to a list of professors. The script uses the `win32com` library to interact with Outlook and send emails. Additionally, it allows for embedding images and attaching files such as flyers and itineraries.

## Features:
- Personalizes emails by inserting the recipient's name in the body of the message.
- Supports the inclusion of an embedded image (e.g., a logo) in the email body.
- Allows for file attachments such as event flyers and itineraries.
- Reads professor names and email addresses from an Excel file.
- Sends emails to each recipient with the specified subject, body, attachments, and image.

## Requirements:
- Python 3.x
- Libraries:
  - `win32com.client` (for Outlook email interaction)
  - `pandas` (for reading Excel files)

  To install the required libraries, use the following commands:
  ```bash
  pip install pywin32 pandas
  ```

## Script Overview:

1. **send_outlook_email**: Sends an email with customizable subject, body, recipient, embedded image, and attachments.
2. **embed_image**: Embeds an image inline in the email body.
3. **attach_file**: Attaches files to the email.
4. **main**: Main function that loads professor data from an Excel file and sends emails to each recipient with the event details.

## Usage:

1. **Prepare the Excel file**:
   - The Excel file (`professors.xlsx`) should have two columns: `Name` and `Email`. Each row should contain the name and email address of a professor.

2. **Customize the script**:
   - Update the file paths for the event flyer, itinerary, and logo image in the script to match your local files.
   - Edit the email body or subject if necessary.

3. **Run the script**:
   - Ensure that you have Outlook installed and configured on your machine.
   - Run the script by executing the following command:
     ```bash
     python send_emails.py
     ```

   The script will loop through the list of professors and send each one a personalized email with the event details, flyer, and itinerary.

## Example Email Body:

The email body is HTML formatted to include the professor's name, event details, and an embedded image:

```html
<html>
  <body>
    <p>Dear Dr. {name},</p>
    <p>We are thrilled to host the <strong>USF ACM AI Convention 2025 on February 8th</strong>, a full-day event featuring keynote speakers...</p>
    <img src="cid:logo_image" alt="ACM Logo" width="300" height="auto" />
  </body>
</html>
```

## Notes:
- Make sure your Outlook application is running and properly configured for sending emails.
- The script assumes the professor's names and email addresses are in an Excel file. You may need to modify the script if your data source is different.
- Ensure that the paths to the files are correctly specified on your machine.

## License:
This script is provided for educational purposes and is free to use or modify.
