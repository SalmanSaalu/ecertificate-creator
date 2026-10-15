
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


BASE_DIR = Path(__file__).resolve().parent.parent
TOKEN_PATH = BASE_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/drive"
]


def get_drive_service():
    if not TOKEN_PATH.exists():
        raise FileNotFoundError(
            f"OAuth token not found: {TOKEN_PATH}. "
            "Run google_auth_setup.py first."
        )

    creds = Credentials.from_authorized_user_file(
        str(TOKEN_PATH), SCOPES
    )

    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            TOKEN_PATH.write_text(
                creds.to_json(), encoding="utf-8"
            )
        else:
            raise RuntimeError(
                "Google credentials are invalid. "
                "Run google_auth_setup.py again."
            )

    return build("drive", "v3", credentials=creds)



import random
import time
from pathlib import Path

from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


def ppt2pdf(input_file_path, nn, folder_id=None):
    input_path = Path(input_file_path)

    if not input_path.is_file():
        raise FileNotFoundError(
            f"PowerPoint file not found: {input_path}"
        )

    service = get_drive_service()

    metadata = {
        "name": Path(nn).stem,
        "mimeType": "application/vnd.google-apps.presentation",
    }

    if folder_id:
        metadata["parents"] = [folder_id]

    for attempt in range(1, 4):
        media = MediaFileUpload(
            str(input_path),
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
            resumable=True,
        )

        try:
            result = (
                service.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id,name,mimeType",
                )
                .execute(num_retries=0)
            )

            presentation_id = result.get("id")

            if not presentation_id:
                raise RuntimeError(
                    f"Google Drive returned no file ID: {result}"
                )

            if result.get("mimeType") != (
                "application/vnd.google-apps.presentation"
            ):
                raise RuntimeError(
                    f"File was not converted to Google Slides: {result}"
                )

            print("Converted presentation:", result.get("name"))
            print("Presentation ID:", presentation_id)

            return presentation_id

        except HttpError as exc:
            status = exc.resp.status
            print(
                f"Google Drive upload attempt {attempt}/3 failed: "
                f"HTTP {status}: {exc.content!r}"
            )

            if status not in (500, 502, 503, 504) or attempt == 3:
                raise

            time.sleep(min(2 ** attempt + random.random(), 10))

    raise RuntimeError("Google Drive upload failed after retries.")



import io

from googleapiclient.http import MediaIoBaseDownload


def export_pdf(presentation_id):
    """Export a Google Slides presentation and return PDF bytes."""

    if not presentation_id or str(presentation_id).strip() == "0":
        raise ValueError(
            f"Invalid Google Slides ID: {presentation_id!r}"
        )

    service = get_drive_service()

    request = service.files().export(
        fileId=str(presentation_id).strip(),
        mimeType="application/pdf",
    )

    pdf_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(pdf_buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    pdf_data = pdf_buffer.getvalue()

    if not pdf_data.startswith(b"%PDF-"):
        raise RuntimeError(
            "Google Drive did not return valid PDF content."
        )

    return pdf_data
