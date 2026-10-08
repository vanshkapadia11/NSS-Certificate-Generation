# 🌾 Grainathon Certificate Automation

Generate personalized certificates from an Excel sheet and email each participant their own PDF—all from a single Python command.

**Excel → Certificate template → Personalized PDFs → Individual emails**

Built for Grainathon events, with a customizable ivory, green, and gold certificate design and Poppins Medium typography. Runs locally without a web interface.

## Features

- **Excel import:** Reads participant names and email addresses from `.xlsx` files.
- **Custom templates:** Supports single-page PDF, PNG, JPG, and JPEG templates.
- **Personalized PDFs:** Adds centered names with configurable position, size, and color. Long names shrink to fit the page.
- **Batch email:** Sends each participant only their own certificate attachment.
- **Duplicate protection:** Tracks SMTP-accepted messages in a local SQLite database.
- **Explicit resending:** Sends updated certificates again when requested with `--resend`.
- **Efficient processing:** Caches template and font preparation, compresses PDFs, and reuses one SMTP connection per batch.
- **Local configuration:** Loads credentials from `.env` and reports progress in the terminal.

## Getting started

### 1. Install dependencies

Install Python 3.10 or newer. The project has been tested with Python 3.12.

From the project directory, run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

On macOS or Linux, activate the environment with `source .venv/bin/activate`.

### 2. Add your participant list

Create `participants.xlsx` in the project directory. Use these headers in the first row:

| Name | Email |
| --- | --- |
| Ananya Sharma | ananya@example.com |
| Aarav Patel | aarav@example.com |

The first worksheet is used by default. Names and email addresses are required. Invalid rows or duplicate email addresses stop the batch before any emails are sent. Completely empty rows are ignored.

Participant spreadsheets are local inputs and are excluded from Git.

### 3. Configure your certificate

Edit `config.json`:

```json
{
  "excel": "participants.xlsx",
  "sheet": null,
  "template": "grainathon-template.pdf",
  "output": "output/certificates",
  "delivery_log": "data/delivery.sqlite3",
  "event_id": "grainathon-2026",
  "organizer": "Grainathon Team",
  "name": {
    "x_percent": 50,
    "y_percent": 52,
    "font_size": 32,
    "color": "#163F34"
  },
  "smtp": {
    "host": "smtp.gmail.com",
    "port": 465,
    "ssl": true,
    "sender": "",
    "username": ""
  }
}
```

| Setting | Purpose |
| --- | --- |
| `excel` | Path to the participant workbook. |
| `sheet` | Worksheet name; `null` selects the first sheet. |
| `template` | Certificate background file with a blank name area. |
| `output` | Directory for generated PDFs. |
| `delivery_log` | Persistent database used to prevent duplicate sends. |
| `event_id` | Unique event identifier. Keep it unchanged when resuming the same event. |
| `organizer` | Organization name used in emails and when rebuilding the bundled design. |
| `name.x_percent` | Horizontal center of the name, as a percentage of page width. |
| `name.y_percent` | Name baseline measured from the top, as a percentage of page height. |
| `name.font_size` | Name size in points. |
| `name.color` | Text color in `#RRGGBB` format. |

File paths are relative to the selected configuration file.

### 4. Generate and review

```powershell
python automate.py
```

Certificates are saved in `output/certificates/`. This command does **not** send email. Review the PDFs and adjust name placement before sending.

## Email setup

Copy the example configuration:

```powershell
Copy-Item .env.example .env
```

Edit `.env` beside `config.json`:

```dotenv
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_SSL=true
SMTP_SENDER=your-email@gmail.com
SMTP_USERNAME=
SMTP_PASSWORD="your-app-password"
```

Leave `SMTP_USERNAME` blank to use `SMTP_SENDER`. Use the credentials required by your provider, including an app password where applicable. The account must support password-based SMTP authentication; OAuth-only accounts are not supported.

| Setting | Behavior |
| --- | --- |
| `SMTP_SSL=true` | Connects using SMTP over SSL. |
| `SMTP_SSL=false` | Connects using SMTP and upgrades with STARTTLS. |
| `SMTP_PORT` | Provider's SMTP port; commonly 465 for SSL or 587 for STARTTLS. |

Existing environment variables override `.env`, which overrides SMTP values in `config.json`. Quoted passwords are supported; `$` characters are treated literally.

**Never commit credentials.** `.env`, participant workbooks, generated certificates, and delivery logs are excluded through `.gitignore`. Keep `.env.example` free of real credentials. If a password has been shared or committed, revoke it and replace it.

## Commands

### Generate certificates only

```powershell
python automate.py
```

### Generate and email certificates

```powershell
python automate.py --send
```

Previously accepted recipients are skipped. New recipients receive their certificates.

### Send updated certificates again

```powershell
python automate.py --send --resend
```

This deliberately sends again to previously accepted recipients in the current spreadsheet and also sends to new recipients. Previous successful records are retained in the `delivery_history` table. Pending or uncertain records remain blocked.

### Use another configuration

```powershell
python automate.py --config path/to/config.json --send
```

The script reads `.env` from the selected configuration's directory. Once configured, it runs without interactive prompts.

## Customize the design

The default `grainathon-template.pdf` uses a warm ivory background, deep green text, gold accents, and grain illustrations. Participant names use bundled **Poppins Medium (weight 500)** for subtle emphasis.

To rebuild the design after editing the organizer in `config.json`:

```powershell
python build_template.py
```

Edit `build_template.py` to change the certificate wording, signature roles, or layout. Rebuilding overwrites `grainathon-template.pdf`.

To use your own design, set `template` in `config.json` to a single-page PDF or supported image, then adjust the name position. Existing template text is not automatically removed. Visually check names using complex scripts before sending.

## Delivery behavior

1. Validate the workbook and render every certificate.
2. Save PDFs locally using stable certificate IDs as filenames.
3. Open and authenticate an SMTP connection when the first message needs sending.
4. Send individual messages and record their status.
5. Stop if a delivery result is uncertain.

A successful run returns exit code `0`; validation errors and blocked or uncertain deliveries return `1`.

**SMTP acceptance means the provider accepted a message, not that it reached the inbox.** Provider limits, network conditions, and attachment sizes affect sending speed.

Keep `data/delivery.sqlite3` between runs. Do not delete it or change the event ID to bypass duplicate protection. A timeout or interrupted run may leave a pending or uncertain record even if the provider accepted the message. Check the provider before manually resolving that record and retrying.

## Troubleshooting

| Message or symptom | What to check |
| --- | --- |
| `skipped (already recorded...)` | Expected duplicate protection. Use `--send --resend` to deliberately send an updated certificate. |
| Missing SMTP settings | Fill in `.env` beside the selected `config.json`, especially sender and password. |
| Incorrect credentials despite editing `.env` | Existing shell environment variables take precedence over `.env`. |
| Pending or uncertain delivery | Check the sender account, SMTP settings, and provider records before retrying. |
| Invalid Excel rows | Correct the reported row numbers, including duplicate emails, and run again. |
| Name is misplaced | Adjust `x_percent`, `y_percent`, or `font_size`, then regenerate without `--send`. |
| Unsupported characters | The bundled font does not cover every script. Review font coverage and rendered output. |
| Slow email delivery | The script reuses its connection, but provider throttling and network latency still apply. |

## Project structure

```text
.
├── automate.py                 # Command-line workflow
├── certificates.py             # Excel, PDF, SMTP, and delivery logging
├── build_template.py           # Grainathon template builder
├── config.json                 # Event and layout configuration
├── .env.example                # Example SMTP settings
├── .gitignore
├── requirements.txt
├── grainathon-template.pdf     # Default certificate background
├── assets/fonts/               # Bundled Poppins fonts and license
└── tests/                      # Validation, rendering, and mocked SMTP tests
```

Local files created or supplied during use include `.env`, `participants.xlsx`, `output/`, and `data/`.

## Tests

```powershell
python -m unittest discover -s tests -v
```

Tests cover Excel validation, PDF text and font rendering, SMTP connection reuse, individual attachments, duplicate protection, resending, and `.env` configuration. Email tests use mocks and do not send real messages.

## Font attribution

Poppins is distributed under the SIL Open Font License. The bundled license is available at [`assets/fonts/OFL.txt`](assets/fonts/OFL.txt). Font source: [Google Fonts — Poppins](https://github.com/google/fonts/tree/main/ofl/poppins).
