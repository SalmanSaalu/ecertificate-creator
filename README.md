# 🎓 E-Certificate Generator

A Django-based web application that automates the generation and distribution of electronic certificates using a PowerPoint template and recipient data from a CSV file.

The application replaces placeholders in a certificate template with individual recipient details, converts the customized presentations into PDF certificates using Google Drive, and is designed to distribute the generated certificates through email.

## 🌐 Live Application

- **Website:** https://ecertificate-creator-1.onrender.com/
- **Upload Documents for reference:** https://ecertificate-creator-1.onrender.com/issuecertificate](https://drive.google.com/drive/folders/1lMAB07UvwiavaC3yNzsCD-OAxaqMvW9e?usp=drive_link
- **GitHub Repository:** https://github.com/SalmanSaalu/ecertificate-creator

> The application is hosted on Render. Availability may vary depending on the service's deployment status and hosting plan.

## ✨ Features

- **Automated Certificate Generation:** Generate personalized certificates from recipient data.
- **PowerPoint Templates:** Use an existing `.pptx` template with placeholders.
- **CSV Integration:** Upload recipient information in CSV format.
- **Dynamic Placeholder Replacement:** Replace template fields with corresponding CSV values.
- **PDF Conversion:** Convert customized PowerPoint presentations into PDF files using Google Drive.
- **Email Integration:** Support certificate delivery through a configured email service.
- **Recipient Validation:** Validate recipient email addresses before processing.
- **Batch Processing:** Process multiple recipients from a single CSV file.
- **Error Handling and Logging:** Track certificate-generation and email-delivery failures.
- **Deployment:** Host the Django application on Render.

## 🛠️ Technology Stack

- Python
- Django
- Pandas
- python-pptx
- Google Drive API
- Google OAuth 2.0
- Gmail API over HTTPS (when configured)
- PostgreSQL / SQLite, depending on environment configuration
- Render
- Git and GitHub

## 🚀 How to Use

### Step 1: Open the application

Visit the [E-Certificate Generator](https://ecertificate-creator-1.onrender.com/).

### Step 2: Open the document upload page

Navigate to the [Certificate Upload Page](https://ecertificate-creator-1.onrender.com/issuecertificate).

### Step 3: Prepare your PowerPoint template

Create a `.pptx` certificate template containing placeholders for dynamic recipient information.

Example:

```text
Certificate of Achievement

This is to certify that <name>
has secured <grade> in <event>
conducted on <date>.
```

Use placeholder names that correspond to the column names in your CSV file. The application's placeholder-mapping logic determines how the values are matched.

### Step 4: Prepare the recipient CSV

Create a CSV file containing recipient information.

Example:

```csv
name,grade,event,date,email
Salman,B,100 Meter,10th June,salman@example.com
Akhil,A,Long Jump,10th June,akhil@example.com
```

Ensure that the CSV contains an `email` column and that recipient email addresses are valid.

### Step 5: Upload both files

Upload:
- The recipient CSV file (`.csv`)
- The PowerPoint certificate template (`.pptx`)

Follow the instructions displayed by the application.

### Step 6: Generate certificates

The application processes the recipient rows, replaces template placeholders, converts the customized presentations to PDF, and attempts email delivery using the configured email integration.

## ⚙️ Run Locally

### Prerequisites

- Python 3.10 or another Python version supported by the project's dependencies
- Git
- A Google Cloud project with the required Google APIs enabled
- Google OAuth credentials for the configured integrations

### 1. Clone the repository

```bash
git clone https://github.com/SalmanSaalu/ecertificate-creator.git
cd ecertificate-creator
```

### 2. Create a virtual environment

```bash
python -m venv venv
```

Activate it on Windows:

```bash
venv\Scripts\activate
```

### 3. Install dependencies

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4. Configure environment variables

Create a `.env` file in the directory containing `manage.py`.

Example configuration:

```dotenv
SECRET_KEY=replace_with_a_secure_secret
DEBUG=True

GOOGLE_TOKEN_FILE=token.json
GOOGLE_DRIVE_FOLDER_ID=your_google_drive_folder_id

# Configure these only if Gmail API sending is implemented.
GMAIL_TOKEN_FILE=gmail_token.json
```

Add any database, allowed-host, email, and other environment variables required by your project's existing Django settings.

Never commit real secrets, OAuth tokens, or private credentials to GitHub.

### 5. Apply database migrations

```bash
python manage.py migrate
```

### 6. Start the development server

```bash
python manage.py runserver
```

Open http://127.0.0.1:8000/ in your browser.

## 🔐 Google API Configuration

### Google Drive integration

The application uses Google Drive to convert PowerPoint presentations into Google Slides presentations and export the results as PDFs.

To configure it:

1. Create or select a project in Google Cloud Console.
2. Enable the Google Drive API.
3. Configure OAuth credentials.
4. Authorize the Google account used by the application.
5. Configure the existing Drive token path and destination folder ID.

Keep the existing Drive token and its permissions separate from the Gmail token.

### Gmail API integration

If Gmail API email delivery is configured:

1. Enable the Gmail API in Google Cloud Console.
2. Request the `https://www.googleapis.com/auth/gmail.send` OAuth scope.
3. Authorize the sending Google account.
4. Store the resulting Gmail token separately from the Drive token.
5. Use the Gmail API over HTTPS to send messages and attach generated PDFs.

The Gmail token should be protected as a secret. OAuth consent-screen settings and Google's token policies may require additional configuration for long-term production use.

## ☁️ Render Deployment

The application can be deployed to Render using the connected GitHub repository.

Typical deployment configuration:

- **Build command:** `pip install -r requirements.txt`
- **Start command:** `gunicorn ecertificate.wsgi:application`

Confirm that the project contains the `gunicorn` dependency and that `ecertificate.wsgi` is the correct WSGI module for your repository.

Configure the necessary environment variables and secret files in Render's dashboard.

Important deployment considerations:

- Set `DEBUG=False` in production.
- Configure `ALLOWED_HOSTS` and trusted CSRF origins for the deployed domain.
- Store API keys, OAuth credentials, and token files securely.
- Do not rely on the local filesystem for permanent user uploads on ephemeral hosting.
- Ensure the production database is configured and migrations are applied.
- Check Render logs when investigating deployment, conversion, or email errors.

## 📁 Important Project Components

| Component | Purpose |
|---|---|
| `home/views.py` | Handles certificate upload, processing, and status reporting |
| `home/convter.py` | Contains PowerPoint conversion and PDF export logic |
| `home/models.py` | Contains application data models, including uploaded-file records |
| `account/` | Account-related application functionality |
| `templates/` | HTML templates used by the web interface |
| `requirements.txt` | Python dependency list |
| `.env` | Local environment configuration; not committed to GitHub |
| `manage.py` | Django management entry point |

The exact contents may vary with the current repository version.

## 🛡️ Security Notes

- Never commit `.env`, OAuth tokens, or client secrets.
- Validate uploaded CSV and PowerPoint files.
- Restrict upload sizes and handle malformed documents safely.
- Protect certificate-generation and status endpoints with appropriate authentication and authorization.
- Keep production secrets in Render environment variables or secret files.
- Follow Google's OAuth policies and email-sending limits.

## 🐛 Troubleshooting

**Application unavailable:** Check the Render service status, deployment history, and logs.

**Google Drive authorization error:** Verify the Drive token, OAuth scopes, and Google API configuration.

**PDF conversion error:** Verify that the Drive API is enabled, the token is valid, and the configured Drive folder is accessible.

**Email sending error:** Check the Gmail token, `gmail.send` scope, OAuth consent configuration, and the Gmail API response.

**CSV processing error:** Confirm that the CSV contains the required columns and that its email addresses and placeholder mappings are valid.

## 📌 Future Improvements

- Downloadable sample CSV and PowerPoint templates
- Certificate preview before generation
- Bulk-processing progress indicator
- Retry handling for transient API failures
- Certificate history and delivery-status dashboard
- Improved upload validation and file-size limits

## 👨‍💻 Author

**SalmanSaalu**

GitHub: https://github.com/SalmanSaalu

---

If you find this project useful, consider starring the repository and sharing feedback or suggestions for improvement.
