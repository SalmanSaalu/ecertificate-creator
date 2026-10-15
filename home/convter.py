import os
import io
import random
import time
from pathlib import Path
import os
import httplib2

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_httplib2 import AuthorizedHttp
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import (
    MediaFileUpload,
    MediaIoBaseDownload,
)


BASE_DIR = Path(__file__).resolve().parent.parent
TOKEN_PATH = Path(
    os.environ.get(
        "GOOGLE_TOKEN_FILE",
        str(BASE_DIR / "token.json"),
    )
)
RUNTIME_TOKEN_PATH = Path("/tmp/token.json")
SCOPES = ["https://www.googleapis.com/auth/drive"]


HTTP_TIMEOUT = 180



def get_drive_service():
    is_render = bool(os.environ.get("RENDER"))

    token_path = TOKEN_PATH

    if is_render:
        if not TOKEN_PATH.exists():
            raise FileNotFoundError(
                f"OAuth token not found: {TOKEN_PATH}. "
                "Check GOOGLE_TOKEN_FILE in Render."
            )

        # Copy the read-only Render secret to writable storage.
        if not RUNTIME_TOKEN_PATH.exists():
            RUNTIME_TOKEN_PATH.parent.mkdir(
                parents=True,
                exist_ok=True,
            )
            RUNTIME_TOKEN_PATH.write_bytes(
                TOKEN_PATH.read_bytes()
            )

        token_path = RUNTIME_TOKEN_PATH

    if not token_path.exists():
        raise FileNotFoundError(
            f"OAuth token not found: {token_path}. "
            "Check your token.json configuration."
        )

    creds = Credentials.from_authorized_user_file(
        str(token_path),
        SCOPES,
    )

    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            token_path.write_text(
                creds.to_json(),
                encoding="utf-8",
            )
        else:
            raise RuntimeError(
                "Google credentials are invalid. "
                "Check your OAuth token and refresh token."
            )

    http = AuthorizedHttp(
        creds,
        http=httplib2.Http(timeout=HTTP_TIMEOUT),
    )

    return build(
        "drive",
        "v3",
        http=http,
        cache_discovery=False,
    )



def ppt2pdf(input_file_path, nn, folder_id=None):
    """
    Upload a PPTX to Google Drive, convert it to Google Slides,
    and return the Google Slides presentation ID.
    """
    input_path = Path(input_file_path)

    if not input_path.is_file():
        raise FileNotFoundError(
            f"PowerPoint file not found: {input_path}"
        )

    if input_path.suffix.lower() != ".pptx":
        raise ValueError(
            "Only .pptx files are supported. "
            "Save the template as PowerPoint Presentation (.pptx)."
        )

    service = get_drive_service()

    metadata = {
        "name": Path(nn).stem,
        "mimeType": "application/vnd.google-apps.presentation",
    }

    if folder_id:
        metadata["parents"] = [folder_id]

    last_error = None

    for attempt in range(1, 4):
        media = None

        try:
            # Use a fresh upload object for every attempt.
            media = MediaFileUpload(
                str(input_path),
                mimetype=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                resumable=True,
                chunksize=1024 * 1024,
            )

            result = (
                service.files()
                .create(
                    body=metadata,
                    media_body=media,
                    fields="id,name,mimeType",
                )
                .execute(num_retries=2)
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

            print(
                "Converted presentation:",
                result.get("name"),
            )
            print("Presentation ID:", presentation_id)

            return presentation_id

        except (HttpError, TimeoutError, OSError) as exc:
            last_error = exc

            status = (
                getattr(exc.resp, "status", None)
                if isinstance(exc, HttpError)
                else None
            )

            # Do not retry permission, authentication, or other
            # permanent HTTP errors.
            retryable = (
                status in (408, 429, 500, 502, 503, 504)
                or isinstance(exc, (TimeoutError, OSError))
            )

            print(
                f"Google Drive upload attempt {attempt}/3 failed: "
                f"{type(exc).__name__}: {exc}"
            )

            if not retryable or attempt == 3:
                raise

            time.sleep(
                min(2 ** attempt + random.random(), 10)
            )

        finally:
            # Close the upload file handle before another attempt
            # or before the caller tries to delete the PPTX.
            if media is not None:
                try:
                    media._fd.close()
                except (AttributeError, OSError):
                    pass

    raise RuntimeError(
        "Google Drive upload failed after retries."
    ) from last_error


def export_pdf(presentation_id):
    """Export Google Slides as PDF bytes."""
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
    downloader = MediaIoBaseDownload(
        pdf_buffer,
        request,
    )

    done = False

    while not done:
        _, done = downloader.next_chunk(num_retries=2)

    pdf_data = pdf_buffer.getvalue()

    if not pdf_data.startswith(b"%PDF-"):
        raise RuntimeError(
            "Google Drive did not return valid PDF content."
        )

    return pdf_data
