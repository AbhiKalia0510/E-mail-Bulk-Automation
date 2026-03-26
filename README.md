# E-mail-Bulk-Automation
Flow automates bulk/individual emails of Available Stock Inventory using Python with Salesforce data. It fetches inventory by Employee Code, creates a pivot-style table, and sends authenticated emails quickly.

## ✅ Linux-ready SMTP version

Files:
- `Bulk_E-Mailing/bulk_email_automation.py`
- `Bulk_E-Mailing/requirements.txt`

## Usage

1. Install dependencies:
   - `python -m pip install -r Bulk_E-Mailing/requirements.txt`

2. Set environment variables:
   - `SF_CLIENT_ID`, `SF_CLIENT_SECRET`, `SF_USERNAME`, `SF_PASSWORD`
   - `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASSWORD`

3. Run:
   - `python Bulk_E-Mailing/bulk_email_automation.py --config path/to/config.xlsx --out path/to/output`

4. Optional:
   - `--smtp-host` and `--smtp-port` can override SMTP default values.

