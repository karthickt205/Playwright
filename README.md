# WhatsApp Web Contact Bot

A Playwright sync automation script that reads contacts from Excel, sends personalized WhatsApp Web messages, captures screenshots, extracts recent messages, and writes JSON and Excel reports.

## Requirements

- Python 3.10 or newer
- Playwright
- openpyxl
- A WhatsApp account
- `contact.xlsx`

Install the Python packages:

```powershell
pip install playwright openpyxl
python -m playwright install chromium
```

## Excel Input

Create `contact.xlsx` in this folder. The first row must contain these headers:

| Name | Phone | Message |
| --- | --- | --- |
| XXXXX | +919876XXXXX | Hello {name} |

`Message` can be empty. The default message is `Hello {name}`. The `{name}` placeholder is replaced with the value from the `Name` column.

Phone numbers may include a plus sign, spaces, or hyphens. They are normalized to digits before WhatsApp search.

## Run

From the workspace root:

```powershell
python playwright-whatsapp-bot\playwright_assign.py
```

Or provide explicit paths:

```powershell
python playwright-whatsapp-bot\playwright_assign.py `
  --contacts "\path\to\contact.xlsx" `
  --output "\path\to\whatsapp_reports" 
 
```

On the first run, WhatsApp Web opens in a visible Chromium window. Scan the QR code manually, then press Enter in the terminal. The persistent browser profile keeps the login for later runs.

## Output

Reports are written to the output directory:

- `whatsapp_report_YYYYMMDD_HHMMSS.json`
- `whatsapp_report_YYYYMMDD_HHMMSS.xlsx`
- `screenshots/` containing a PNG for each successfully sent message

Each report row includes the name, masked phone number, personalized message, sent status, up to three extracted messages, screenshot path, error details, and processing time. Only the last four phone digits are visible in reports.

## Error Handling

A failure for one contact is recorded in the report and does not stop the remaining contacts. Check the `Error` column or JSON field for details.

If the workbook is not found, pass its full path with `--contacts`. If WhatsApp asks for login again, scan the QR code and reuse the same `--browser-data` directory.
