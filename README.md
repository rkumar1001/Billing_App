# Billing App — Invoice Duplicator

A cross-platform Python desktop app for duplicating monthly invoices. You point it at an existing Excel breakdown + Word invoice pair, map each editable field to a cell (click-through wizard), and from then on regenerating next month's invoice is an edit-and-click operation. Targets Windows, macOS, Linux.

## Features

- **Bill Profiles** — save an Excel + Word template pair with a field-to-coordinate map.
- **One-click regeneration** — pick month, hit Generate; the app copies the templates, fills the dates column for every day of the month, updates the invoice number, date, billing period, total hours, grand total, and saves to the folder you choose.
- **Atomic invoice numbering** per profile — no duplicates even under concurrent clicks.
- **Multi-client** — unlimited profiles.
- **Preserves template formatting** — images, fonts, formulas stay intact (Excel recomputes formulas on open).

## Install (dev)

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python -m billing_app.main
```

## First run

1. Click **+ New Profile**.
2. *Step 1 — Files*: name the profile, pick your existing `.xlsx` and `.docx`, set invoice prefix and the next invoice number.
3. *Step 2 — Excel Mapping*: for each field (invoice date, month number, dates column start, total hours, grand total), enter the cell address (e.g. `B5`). The preview on the right shows the sheet so you can find cells.
4. *Step 3 — Word Mapping*: for each field (invoice number, invoice date, billing period, total hours, grand total) enter table / row / col / paragraph index. The right pane shows every table and paragraph in the document.
5. *Step 4 — Review*: confirm and save.

Then from the Dashboard, click **Generate** on your profile, pick the month/year, review the computed totals, and click Generate.

## Building a standalone executable

```bash
pip install pyinstaller
pyinstaller --clean --noconfirm build.spec
```

Outputs:

- Windows: `dist/BillingApp.exe`
- macOS: `dist/BillingApp.app`
- Linux: `dist/BillingApp`

For Windows, consider wrapping `dist/BillingApp.exe` with Inno Setup for a proper installer and code-signing to avoid SmartScreen warnings.

## Data locations

- SQLite database and generated invoices: user data dir (e.g. `~/.local/share/BillingApp/` on Linux, `%APPDATA%\BillingApp` on Windows, `~/Library/Application Support/BillingApp` on macOS).
- Settings JSON: user config dir.

## Roadmap

- Click-to-map cells in Excel preview (currently: address entry).
- Quick Mode for ad-hoc invoices without creating a profile.
- PDF export via `docx2pdf` (Windows) or LibreOffice headless (mac/Linux).
- SMTP email delivery with credentials stored in the OS keyring.
- CI matrix builds for all three OSes on tag push.
