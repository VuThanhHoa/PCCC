from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ['https://www.googleapis.com/auth/drive.file']

flow = InstalledAppFlow.from_client_secrets_file(
    'path/to/credentials.json', SCOPES)
credentials = flow.run_local_server(port=0)

# Save the credentials for the next run
with open('token.json', 'w') as token:
    token.write(credentials.to_json())
