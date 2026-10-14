
import re
import logging
from pathlib import Path
import os
from dotenv import load_dotenv
load_dotenv()
import pandas as pd
import requests

from pptx import Presentation

from django.conf import settings
from django.contrib import messages
from django.core.mail import EmailMessage
from django.shortcuts import render, redirect
from django.utils.crypto import get_random_string
from django.views import View

from .models import beforefetch
from .convter import ppt2pdf, export_pdf
from home.models import statusdetails


logger = logging.getLogger(__name__)

BASE_DIR = Path(settings.BASE_DIR)
MEDIA_DIR = BASE_DIR / "media"
OUTPUT_DIR = MEDIA_DIR / "after_fetch"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def first(request):
    return render(request, "index.html")


def get_placeholders(presentation):
    """Find placeholders such as <Name> and <Email> in a PPTX."""
    placeholders = []

    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for match in re.findall(r"<[^<>]+>", shape.text):
                if match not in placeholders:
                    placeholders.append(match)

    return placeholders


def replace_placeholders(presentation, replacements):
    """
    Replace placeholders in slide text.

    Setting paragraph.text may change the original run-level formatting
    within that paragraph, but avoids failures when a placeholder is
    split across multiple PowerPoint text runs.
    """
    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            for paragraph in shape.text_frame.paragraphs:
                original_text = paragraph.text
                updated_text = original_text

                for placeholder, value in replacements.items():
                    updated_text = updated_text.replace(
                        placeholder, str(value)
                    )

                if updated_text != original_text:
                    paragraph.text = updated_text




class issueCertificate(View):

    def get(self, request):
        return render(request, "issuecertificate.html")

    def post(self, request):
        uploaded_csv = request.FILES.get("excel")
        uploaded_template = request.FILES.get("template")

        if not uploaded_csv or not uploaded_template:
            messages.error(
                request,
                "Please upload both a CSV file and a PowerPoint template."
            )
            return render(request, "issuecertificate.html")

        if not uploaded_csv.name.lower().endswith(".csv"):
            messages.error(request, "Please upload a valid CSV file.")
            return render(request, "issuecertificate.html")

        if not uploaded_template.name.lower().endswith(
            (".ppt", ".pptx")
        ):
            messages.error(
                request,
                "Please upload a PowerPoint template."
            )
            return render(request, "issuecertificate.html")

        try:
            df = pd.read_csv(uploaded_csv)

            if df.empty:
                messages.error(request, "The CSV file contains no data.")
                return render(request, "issuecertificate.html")

            # Your existing template uses forexcel for its column choices.
            forexcel = [str(column) for column in df.columns]

            # python-pptx supports PPTX, not legacy binary PPT files.
            if not uploaded_template.name.lower().endswith(".pptx"):
                messages.error(
                    request,
                    "Please save your PowerPoint template as .pptx "
                    "and upload it again."
                )
                return render(request, "issuecertificate.html")

            presentation = Presentation(uploaded_template)
            fortemplate = get_placeholders(presentation)

            if not fortemplate:
                messages.error(
                    request,
                    "No placeholders were found. Use placeholders "
                    "such as <Name> or <Course> in the template."
                )
                return render(request, "issuecertificate.html")

            unique_id = get_random_string(length=32)

            uploaded_csv.seek(0)
            uploaded_template.seek(0)

            beforefetch.objects.create(
                excel_file_upload=uploaded_csv,
                template_file_upload=uploaded_template,
                uniquename=unique_id
            )

            return render(
                request,
                "certificateparttwo.html",
                {
                    "forexcel": forexcel,
                    "fortemplate": fortemplate,
                    "de": unique_id,
                }
            )

        except Exception:
            logger.exception("Error preparing certificate template.")
            messages.error(
                request,
                "Could not read the CSV or PowerPoint template. "
                "Check the server log for details."
            )
            return render(request, "issuecertificate.html")


class Certificateparttwo(View):

    def get(self, request):
        return redirect("/")

    def post(self, request):
        event_id = request.POST.get("event")

        if not event_id:
            messages.error(request, "Missing certificate request ID.")
            return redirect("/")

        try:
            upload = beforefetch.objects.get(uniquename=event_id)

            # Use the saved file paths, not the original upload objects.
            with upload.excel_file_upload.open("rb") as csv_file:
                df = pd.read_csv(csv_file)

            if df.empty:
                messages.error(request, "The CSV file contains no data.")
                return redirect("/")

            template_path = upload.template_file_upload.path

            # Load the original presentation for placeholder discovery.
            original_presentation = Presentation(template_path)
            placeholders = get_placeholders(original_presentation)

            if not placeholders:
                messages.error(
                    request,
                    "No placeholders were found in the PowerPoint template."
                )
                return redirect("/")

            # Match each placeholder to the column selected in the form.
            replacements_by_placeholder = {}

            for index, placeholder in enumerate(placeholders, start=1):
                selected_column = request.POST.get(str(index))

                if selected_column not in [
                    str(column) for column in df.columns
                ]:
                    messages.error(
                        request,
                        f"Invalid or missing CSV column selected "
                        f"for {placeholder}."
                    )
                    return redirect("/")

                replacements_by_placeholder[placeholder] = selected_column

            sent_count = 0
            failed_count = 0

            # Generate one certificate for each CSV row.
            for row_number, (_, row) in enumerate(df.iterrows(), start=1):
                recipient = None

                # Find the first plausible email address in the row.
                for cell in row:
                    candidate = str(cell).strip()

                    if re.fullmatch(
                        r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+"
                        r"@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+",
                        candidate
                    ):
                        recipient = candidate
                        break

                if not recipient:
                    logger.warning(
                        "Row %s has no email address; skipping.",
                        row_number
                    )
                    failed_count += 1
                    continue

                pdf_path = None
                pptx_path = None

                try:
                    # Use a fresh presentation for every recipient.
                    presentation = Presentation(template_path)

                    replacements = {}

                    for placeholder, column in (
                        replacements_by_placeholder.items()
                    ):
                        value = row[column]

                        if pd.isna(value):
                            value = ""

                        replacements[placeholder] = value

                    replace_placeholders(presentation, replacements)

                    # Unique filenames prevent certificates from
                    # overwriting one another.
                    filename = get_random_string(20)

                    pptx_path = OUTPUT_DIR / f"{filename}.pptx"
                    pdf_path = OUTPUT_DIR / f"{filename}.pdf"

                    presentation.save(str(pptx_path))

                    # This existing helper is expected to upload/convert
                    # the PPTX and return a Google Slides presentation ID.
                    presentation_id = ppt2pdf(
                        str(pptx_path),
                        pptx_path.name,
                        folder_id=os.getenv("GOOGLE_DRIVE_FOLDER_ID"),
                    )

                    pdf_data = export_pdf(presentation_id)


                    pdf_path.write_bytes(pdf_data)

                    # Confirm that the saved attachment is also a PDF.
                    if pdf_path.read_bytes()[:5] != b"%PDF-":
                        raise ValueError(
                            "The saved certificate is not a valid PDF."
                        )

                    email = EmailMessage(
                        subject="Your Certificate",
                        body=(
                            "Congratulations! Please find your "
                            "certificate attached."
                        ),
                        from_email=settings.DEFAULT_FROM_EMAIL,
                        to=[recipient],
                    )

                    email.attach(
                        filename="certificate.pdf",
                        content=pdf_data,
                        mimetype="application/pdf",
                    )

                    # Send only after PDF validation succeeds.
                    result = email.send(fail_silently=False)

                    if result != 1:
                        raise RuntimeError(
                            "Django did not confirm that the email was sent."
                        )

                    statusdetails.objects.create(
                        email=recipient,
                        statusnamecheck=request.user.id,
                        send=1
                    )

                    sent_count += 1

                except Exception:
                    logger.exception(
                        "Certificate generation/email failed for CSV row %s.",
                        row_number
                    )

                    failed_count += 1

                    # Record the failure when the model permits it.
                    try:
                        statusdetails.objects.create(
                            email=recipient,
                            statusnamecheck=request.user.id,
                            send=0
                        )
                    except Exception:
                        logger.exception(
                            "Could not save failure status for %s.",
                            recipient
                        )

                finally:
                    # Remove temporary files without affecting
                    # the email attachment already held in memory.
                    for path in (pptx_path, pdf_path):
                        if path and path.exists():
                            try:
                                path.unlink()
                            except OSError:
                                logger.exception(
                                    "Could not remove temporary file %s",
                                    path
                                )

            messages.success(
                request,
                f"Certificate processing finished. "
                f"Emails sent: {sent_count}. "
                f"Failed or skipped: {failed_count}."
            )

        except beforefetch.DoesNotExist:
            messages.error(
                request,
                "The certificate request was not found. Please start again."
            )

        except Exception:
            logger.exception("Certificate processing failed.")
            messages.error(
                request,
                "Certificate processing failed. Check the server logs "
                "for the detailed error."
            )

        return redirect("/")
        

def status(request):
    ee = request.user.id
    ff = statusdetails.objects.filter(
        statusnamecheck=ee
    ).order_by("-dateadded")

    return render(
        request,
        "status.html",
        {"statusdetails": ff}
    )
