# MLSA Certificate Sender

Safely match PDF certificates to participants in an Excel workbook and send them through Gmail.

The script is conservative by design:

- A normal run is a dry run and never sends email.
- Matching uses the participant name and the PDF filename.
- Uncertain matches, duplicate records, invalid email addresses, and duplicate email assignments are not sent.
- A test mode sends one selected certificate only to a test address.
- Live sending requires an explicit `SEND` confirmation.
- Successful sends are recorded so the same recipient/PDF pair is not sent again.

## Requirements

- Windows PowerShell
- Python 3.10 or newer
- A Gmail account with 2-Step Verification enabled
- A Gmail App Password for the configured sender account

The sender address is currently configured in `send_certificates.py` as `mlsadypcet@gmail.com`.

## Project files

| File or directory | Purpose |
| --- | --- |
| `send_certificates.py` | Matches certificates and optionally sends email |
| `students.xlsx` | Participant workbook |
| `certificates.zip` | Source archive containing PDF certificates |
| `certificates/` | Extracted PDFs, created automatically when needed |
| `requirements.txt` | Python dependencies |
| `report.csv` | Latest matching and sending results |
| `send_log.csv` | Persistent record of submitted emails |

The workbook must contain these columns:

- `Full name`
- `Email2`

## Setup

Open PowerShell in the project directory and run:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The script extracts `certificates.zip` into `certificates/` automatically if no PDFs are present. Keep the ZIP archive and workbook in the same directory as the script.

## Gmail setup

1. Enable 2-Step Verification for the sender Gmail account.
2. Create an App Password at <https://myaccount.google.com/apppasswords>.
3. Keep the App Password private. Never add it to this repository or commit it to source control.

The script prompts for the App Password securely when test or live sending begins. Do not use the regular Gmail password.

## Usage

### 1. Run a dry run

```powershell
.\.venv\Scripts\python.exe send_certificates.py
```

Review `report.csv`. Safe matches have one of these statuses:

- `MATCHED`
- `MATCHED_TOP_PERFORMER`

Resolve all `AMBIGUOUS_OR_NO_MATCH`, `DUPLICATE_EMAIL`, `MISSING_NAME`, and `MISSING_OR_INVALID_EMAIL` records before sending.

### 2. Send one test email

Send the first safe match to your own inbox:

```powershell
.\.venv\Scripts\python.exe send_certificates.py --test-email your-email@example.com
```

Select a specific participant when needed:

```powershell
.\.venv\Scripts\python.exe send_certificates.py `
  --test-email your-email@example.com `
  --test-name "Sujal Vishwakarma"
```

Test mode always sends to the address supplied with `--test-email`, never to the participant.

### 3. Send certificates live

After reviewing the report and test email:

```powershell
.\.venv\Scripts\python.exe send_certificates.py --send
```

The script stops if unresolved records remain. If the records are safe, type `SEND` when prompted. The script skips recipient/PDF pairs already recorded as `SENT` in `send_log.csv`.

## Certificate matching

Matching is based on normalized names and fuzzy filename comparison. A match must meet the configured score and confidence-gap thresholds in `send_certificates.py`.

PDFs containing `Top Performer` in the filename are preferred for their participant and replace that participant's standard certificate. Confirm these assignments in `report.csv` before live sending.

## Safety and recovery

- Keep `send_log.csv` between runs; deleting it removes the duplicate-send protection.
- SMTP acceptance means Gmail accepted the message, not that it reached the inbox.
- If the connection fails after Gmail accepts a message but before the log is updated, check Gmail Sent before retrying.
- Review `report.csv` after every run.
- Do not commit App Passwords, private participant data, or generated reports unless the project requires them.

## Troubleshooting

### Required workbook column is missing

Make sure the first worksheet contains exactly `Full name` and `Email2` column headers.

### No PDFs are found

Confirm that `certificates.zip` is beside `send_certificates.py`, or place the extracted PDFs under `certificates/`.

### Gmail login fails

Confirm that 2-Step Verification is enabled and that you are using a Gmail App Password rather than the regular account password.

### A participant is not matched

Check the PDF filename and participant name in `report.csv`. Rename the PDF or correct the workbook entry only after manually verifying the identity.

## Dependencies

- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [RapidFuzz](https://rapidfuzz.github.io/RapidFuzz/)

