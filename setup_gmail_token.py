from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "gmail_credential.json"
TOKEN_FILE = BASE_DIR / "gmail_token.json"


def main():
    if not CREDENTIALS_FILE.exists():
        raise FileNotFoundError(
            f"Missing OAuth credentials: {CREDENTIALS_FILE}"
        )

    flow = InstalledAppFlow.from_client_secrets_file(
        str(CREDENTIALS_FILE),
        SCOPES,
    )

    credentials = flow.run_local_server(
        host="localhost",
        port=0,
        open_browser=True,
    )

    TOKEN_FILE.write_text(credentials.to_json(), encoding="utf-8")

    print("Gmail authorization completed.")
    print(f"Token saved to: {TOKEN_FILE}")
    print("Keep this token private. Do not upload it to GitHub.")


if __name__ == "__main__":
    main()