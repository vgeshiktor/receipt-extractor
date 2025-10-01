from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
import os

SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/drive.file'
]

creds = None
if os.path.exists('credentials'):
    with open('credentials', 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'client_secret.json', SCOPES
        )
        # This line is crucial! It requests a refresh token.
        creds = flow.run_local_server(host="localhost", port=8080)

    with open('credentials', 'wb') as token:
        pickle.dump(creds, token)