# Brevo Job Mailer

Production-ready CustomTkinter desktop app for sending personalized job application emails to HR contacts through Brevo SMTP.

## Features

- Reads the `HR Contacts` sheet from the configured `.xlsx` file under `excel/`.
- Supports `.csv` for small local tests with the same columns.
- Required contact fields: `Name`, `Company`, `Email`.
- Extracts `{{FIRST_NAME}}` from `Name`; blank names become `Team`.
- Sends HTML email from `templates/email.html` with a plain-text fallback.
- Attaches `resume/Ganesh_Mishra_Resume.pdf` by default and allows changing the PDF in the UI.
- Skips empty, invalid, duplicate, and previously sent emails.
- Maintains `sent_emails.json` so the same email is never sent twice.
- Waits 60 seconds between emails, uses 3 retries, and automatic SMTP reconnect waiting.
- Sends only between `08:00` and `12:00`, with a daily cap of `300` emails.
- Provides Start, Pause, Resume, Stop, Search HR, Preview Email, Preview Resume, Retry Failed Emails, Schedule Sending, and Export Logs.
- Generates `sent_report.xlsx`, `failed_report.xlsx`, and `skipped_report.xlsx`.
- Writes `logs/sent.log` and `logs/failed.log`.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Update `.env` with your real Brevo SMTP key:

```env
BREVO_SMTP_SERVER=smtp-relay.brevo.com
BREVO_SMTP_PORT=587
BREVO_SMTP_LOGIN=b15102001@smtp-brevo.com
BREVO_SMTP_KEY=your-real-brevo-smtp-key
SENDER_EMAIL=ganeshkemoap123@gmail.com
SENDER_NAME=Ganesh Mishra
```

`SENDER_EMAIL` must be a verified sender/domain email in your Brevo account.

Do not commit `.env`. It is ignored by Git. Use `.env.example` as the template.

## Run

```powershell
& "C:\Users\Ezhil\AppData\Local\Programs\Python\Python312\python.exe" main.py
```

You can also double-click `run_app.bat` or run `.\run_app.ps1` from PowerShell.

Place your real files here before running:

```text
excel/kaamkibaatein_HR_Contact_Database.xlsx
resume/Ganesh_Mishra_Resume.pdf
```

Actual `.xlsx` and `.pdf` files are ignored by Git so private HR data and resumes are not uploaded. The folders are kept with `.gitkeep`.

## Excel Format

The app always reads the sheet named `HR Contacts`.

Expected columns:

| S.No | Name | Title | Company | Category | Email |
| --- | --- | --- | --- | --- | --- |
| 1 | Reena Vijayanand | HR Manager | AdaniConneX | Cloud | reena.v@adani.com |
| 2 | Puja Gupta | Talent Acquisition | Affle | Product | puja@affle.com |

Only `Name`, `Company`, and `Email` are required by the app. Other columns are preserved for future personalization.

## Config

`config.json`:

```json
{
  "daily_limit": 300,
  "send_window_enabled": true,
  "send_window_start": "08:00",
  "send_window_end": "12:00",
  "min_delay": 60,
  "max_delay": 60,
  "retry": 3
}
```

With this config, if the app is started outside `08:00-12:00`, it waits automatically. After 300 successful sends in a day, it waits until the next day's 08:00 window and continues with the remaining contacts. The actual config file also stores paths, template name, logs/reports folders, and optional schedule time.

Default file paths:

```json
{
  "excel_path": "excel/kaamkibaatein_HR_Contact_Database.xlsx",
  "resume_path": "resume/Ganesh_Mishra_Resume.pdf"
}
```

## Personalization

Edit `templates/email.html`.

Available placeholders:

- `{{FIRST_NAME}}`
- `{{NAME}}`
- `{{COMPANY}}`
- `{{EMAIL}}`
- `{{TITLE}}`
- `{{CATEGORY}}`

Default subject:

```text
Application for Backend / Full Stack Developer | Ganesh Mishra
```

## Reports and Logs

Reports are written into `reports/`:

- `sent_report.xlsx`
- `failed_report.xlsx`
- `skipped_report.xlsx`

Each report contains:

```text
Name, Company, Email, Timestamp, Status, Reason
```

Logs are written into `logs/`:

- `sent.log`
- `failed.log`

## Safety Notes

- Start with a small test list before sending to all 1800 contacts.
- The app is configured for 300 emails per day, only between 08:00 and 12:00.
- If internet drops, the worker waits for SMTP connectivity and resumes automatically.
- If the app closes, restart with the same workbook. Emails in `sent_emails.json` are skipped.
- Do not commit or share a real `BREVO_SMTP_KEY`.

## PyInstaller Build

```powershell
pyinstaller --noconfirm --onefile --windowed --name BrevoJobMailer --add-data "templates;templates" --add-data "config.json;." --add-data ".env;." --add-data "resume;resume" main.py
```

Debug-friendly folder build:

```powershell
pyinstaller --noconfirm --onedir --windowed --name BrevoJobMailer --add-data "templates;templates" --add-data "config.json;." --add-data ".env;." --add-data "resume;resume" main.py
```
