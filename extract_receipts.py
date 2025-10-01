import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
import base64


# Define the scopes your application needs
SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.file'
]

# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret.json', SCOPES)
        creds = flow.run_local_server(host="localhost", port=8080)

    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

# Gmail and Google Drive setup
gmail = build('gmail', 'v1', credentials=creds)
drive = build('drive', 'v3', credentials=creds)
folder_id = '1eCFGIbK4jgwoGtGfruzhcTBxq-gwWZBq'  # Replace with your folder ID

# Fetch emails from date
query = 'has:attachment after:2025/07/01 before:2025/08/01'
results = gmail.users().messages().list(userId='me', q=query).execute()
messages = results.get('messages', [])

os.makedirs('downloaded_receipts', exist_ok=True)

for msg in messages:
    email = gmail.users().messages().get(userId='me', id=msg['id']).execute()
    parts = email.get('payload', {}).get('parts', [])
    for part in parts:
        if part['filename']:
            attachment = gmail.users().messages().attachments().get(
                userId='me', messageId=msg['id'], id=part['body']['attachmentId']
            ).execute()

            file_data = base64.urlsafe_b64decode(attachment['data'])
            path = f"downloaded_receipts/{part['filename']}"

            with open(path, 'wb') as f:
                f.write(file_data)

            # Upload to Google Drive
            file_metadata = {'name': part['filename'], 'parents': [folder_id]}

            # Create a MediaFileUpload object from the local file path
            media = MediaFileUpload(path, mimetype='application/octet-stream')

            drive.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()

print("Attachments uploaded successfully!")
