import imaplib
import email
# from email.header import decode_header
import os

# account credentials
username = 'johndoe@gmail.com' # GANTI DENGAN EMAIL ANDA
password = 'johndoe1234' # GANTI DENGAN PASSWORD ANDA
imap_url = 'imap.gmail.com'
attachment_dir = 'C:/Users/USER/Documents/Github/email-automation/files' # GANTI DENGAN LOKASI FOLDER YANG ANDA INGINKAN


# connect to the email server
mail = imaplib.IMAP4_SSL(imap_url)
mail.login(username, password)
mail.select("inbox") # select the inbox

# search for specific emails by subject
result, data = mail.search(None, 'SUBJECT "Data_Report"')
email_ids = data[0].split()
# Check if we found any emails
if email_ids:
    # Select the most recent email ID
    latest_email_id = email_ids[-1]
    
    # Fetch the email by ID
    typ, msg_data = mail.fetch(latest_email_id, '(RFC822)')
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if filename:
                    filepath = os.path.join(attachment_dir, filename)
                    with open(filepath, 'wb') as fp:
                        fp.write(part.get_payload(decode=True))
                    print(f'Downloaded "{filename}" to "{filepath}"')
else:
    print("No emails found with the specified subject.")

# Log out
mail.logout()