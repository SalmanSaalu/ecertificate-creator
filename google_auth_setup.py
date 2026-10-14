
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

BASE_DIR = Path(__file__).resolve().parent

SCOPES = ["https://www.googleapis.com/auth/drive"]

credentials_path = BASE_DIR / "credentials.json"
token_path = BASE_DIR / "token.json"

creds = None

if token_path.exists():
    creds = Credentials.from_authorized_user_file(
        str(token_path), SCOPES
    )

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(credentials_path), SCOPES
        )
        creds = flow.run_local_server(port=0)

    token_path.write_text(creds.to_json(), encoding="utf-8")

print("Google Drive authorization completed.")
print("Token saved to:", token_path)
