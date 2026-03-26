# -*- coding: utf-8 -*-
"""
Bulk email automation for Salesforce inventory.

Designed for cross-platform use with SMTP (Linux friendly).
Requires Salesforce credentials and SMTP credentials in environment variables.
"""

import os
import sys
import logging
from datetime import datetime
import argparse
import requests
import pandas as pd
from pathlib import Path
from email.message import EmailMessage
import smtplib

# Constants
MAX_ATTACHMENT_MB = 30
DEFAULT_OUTPUT_FOLDER = "./output"

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def load_config_sheet(config_file: Path) -> pd.DataFrame:
    df = pd.read_excel(config_file)
    df.columns = df.columns.str.strip()
    return df


def salesforce_auth():
    client_id = os.getenv("SF_CLIENT_ID")
    client_secret = os.getenv("SF_CLIENT_SECRET")
    username = os.getenv("SF_USERNAME")
    password = os.getenv("SF_PASSWORD")

    if not all([client_id, client_secret, username, password]):
        raise ValueError(
            "Missing one or more Salesforce credentials in env vars: SF_CLIENT_ID, SF_CLIENT_SECRET, SF_USERNAME, SF_PASSWORD"
        )

    auth_url = (
        "https://login.salesforce.com/services/oauth2/token"
        f"?grant_type=password&client_id={client_id}"
        f"&client_secret={client_secret}"
        f"&username={username}&password={password}"
    )

    resp = requests.post(auth_url)
    resp.raise_for_status()
    data = resp.json()

    return data["access_token"], data["instance_url"]


def build_pivot_html(df: pd.DataFrame, to_email: str) -> str:
    location = df["Location.Name"].iloc[0] if "Location.Name" in df.columns else "Unknown"

    order = [
        "Terminal(POS)", "Material", "Adapter", "Cable", "Battery",
        "Biometric", "Soundbox", "Stand", "Base", "Paper POS",
        "SIM", "Tool Kit", "Paper Rolls", "POSM Kit Stickers",
        "Back Cover", "Sticker", "Pendrive", "POS Stickers", "Tent Card"
    ]

    df["Product_Type__c"] = df.get("Product_Type__c", pd.Series(dtype=str)).fillna("Others")

    html = f"""
    <p style=\"font-family:Calibri;font-size:14px;\">\n    Dear {to_email},<br><br>\n    Please find the stock inventory with <b>{location}</b>.<br><br>\n    </p>\n
    <table style=\"border-collapse:collapse;font-family:Calibri;font-size:14px;width:900px;\">\n    <tr style=\"background-color:#B7DEE8;font-weight:bold;border-bottom:3px solid #1F4E79;\">\n        <td colspan=\"3\" style=\"padding:10px;\">\n            <b>Location: Location Name</b> &nbsp;&nbsp;&nbsp;&nbsp; {location}\n        </td>\n    </tr>\n
    <tr style=\"background-color:#D9E1F2;font-weight:bold;\">\n        <th style=\"border:1px solid #000;padding:8px;\">Product Type</th>\n        <th style=\"border:1px solid #000;padding:8px;\">Product Name</th>\n        <th style=\"border:1px solid #000;padding:8px;text-align:right;\">Quantity</th>\n    </tr>\n    """

    grand_total = 0

    for category in order:
        group = df[df["Product_Type__c"] == category]
        if group.empty:
            continue

        group = group.sort_values(by="Product2.Name") if "Product2.Name" in group.columns else group
        category_total = 0

        html += f"""
        <tr style=\"background:#BFBFBF;font-weight:bold;\">\n            <td style=\"border:1px solid #000;\">{category}</td>\n            <td style=\"border:1px solid #000;\"></td>\n            <td style=\"border:1px solid #000;\"></td>\n        </tr>\n        """

        for _, row in group.iterrows():
            qty = int(row.get("QuantityOnHand", 0) or 0)
            category_total += qty

            product_name = row.get("Product2.Name", "")
            html += f"""
            <tr>\n                <td style=\"border-left:1px solid #000;border-right:1px solid #000;font-weight:bold;\">\n                    {category}\n                </td>\n                <td style=\"border-right:1px solid #000;\">\n                    {product_name}\n                </td>\n                <td style=\"border-right:1px solid #000;text-align:right;\">\n                    {qty}\n                </td>\n            </tr>\n            """

        html += f"""
        <tr style=\"background:#6F6FAE;color:white;font-weight:bold;\">\n            <td style=\"border:1px solid #000;\">{category} Total</td>\n            <td style=\"border:1px solid #000;\"></td>\n            <td style=\"border:1px solid #000;text-align:right;\">{category_total}</td>\n        </tr>\n        """

        html += "<tr><td colspan='3' style='height:10px;'></td></tr>"
        grand_total += category_total

    html += f"""
    <tr style=\"background:#A6A6A6;font-weight:bold;\">\n        <td style=\"border:1px solid #000;\">Grand Total</td>\n        <td style=\"border:1px solid #000;\"></td>\n        <td style=\"border:1px solid #000;text-align:right;\">{grand_total}</td>\n    </tr>\n    </table>\n    """

    return html


def fetch_inventory(instance_url: str, access_token: str, location_code: str) -> pd.DataFrame:
    query = f"""
    SELECT Location.Emp_Code__c, Location.Name,
           Product_Type__c, Product2.Name, QuantityOnHand
    FROM ProductItem
    WHERE Location.Emp_Code__c = '{location_code}' AND QuantityOnHand > 0
    """

    url = f"{instance_url}/services/data/v58.0/query"
    headers = {"Authorization": f"Bearer {access_token}"}

    res = requests.get(url, headers=headers, params={"q": query})
    res.raise_for_status()
    records = res.json().get("records", [])

    return pd.json_normalize(records) if records else pd.DataFrame()


def send_email(smtp_host: str, smtp_port: int, smtp_user: str, smtp_password: str,
               to_email: str, cc_email: str, subject: str,
               html_body: str, attachment_path: Path | None = None):
    msg = EmailMessage()
    msg["From"] = smtp_user
    msg["To"] = to_email
    if cc_email:
        msg["Cc"] = cc_email
    msg["Subject"] = subject
    msg.set_content("This email contains HTML content.")
    msg.add_alternative(html_body, subtype="html")

    if attachment_path and attachment_path.exists():
        size_mb = attachment_path.stat().st_size / (1024 * 1024)
        if size_mb <= MAX_ATTACHMENT_MB:
            with open(attachment_path, "rb") as f:
                data = f.read()

            msg.add_attachment(
                data,
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=attachment_path.name,
            )
        else:
            logging.warning("Attachment %s is %.2f MB > %d MB, skipping attachment.", attachment_path, size_mb, MAX_ATTACHMENT_MB)

    with smtplib.SMTP(smtp_host, smtp_port) as smtp:
        smtp.starttls()
        smtp.login(smtp_user, smtp_password)
        smtp.send_message(msg)


def main():
    parser = argparse.ArgumentParser(description="Bulk mail automation for Salesforce inventory.")
    parser.add_argument("--config", required=True, help="Configuration Excel file path")
    parser.add_argument("--out", default=DEFAULT_OUTPUT_FOLDER, help="Output directory")
    parser.add_argument("--smtp-host", default=os.getenv("SMTP_HOST", "smtp.gmail.com"), help="SMTP server host")
    parser.add_argument("--smtp-port", type=int, default=int(os.getenv("SMTP_PORT", 587)), help="SMTP server port")
    args = parser.parse_args()

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    smtp_user = os.getenv("SMTP_USER")
    smtp_password = os.getenv("SMTP_PASSWORD")
    if not (smtp_user and smtp_password):
        raise ValueError("SMTP_USER and SMTP_PASSWORD must be set in environment variables.")

    access_token, instance_url = salesforce_auth()
    config_df = load_config_sheet(Path(args.config))

    for _, cfg in config_df.iterrows():
        loc_code = str(cfg.get("Location.Emp_Code__c", "")).strip()
        to_email = str(cfg.get("TO_EMAIL", "")).strip()
        cc_email = str(cfg.get("CC_EMAIL", "")).strip()
        subject = str(cfg.get("SUBJECT", f"Stock Report - {loc_code}"))

        if not loc_code or not to_email:
            logging.warning("Skipping row with missing Location.Emp_Code__c or TO_EMAIL: %s", cfg)
            continue

        logging.info("Processing location %s -> %s", loc_code, to_email)

        try:
            df = fetch_inventory(instance_url, access_token, loc_code)
            if df.empty:
                logging.info("No inventory found for %s", loc_code)
                continue

            html_body = build_pivot_html(df, to_email)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = out_dir / f"Stock_{loc_code}_{timestamp}.xlsx"
            df.to_excel(file_path, index=False)

            send_email(
                smtp_host=args.smtp_host,
                smtp_port=args.smtp_port,
                smtp_user=smtp_user,
                smtp_password=smtp_password,
                to_email=to_email,
                cc_email=cc_email,
                subject=subject,
                html_body=html_body,
                attachment_path=file_path,
            )

            logging.info("Email sent successfully for %s", loc_code)

        except Exception as exc:
            logging.exception("Failed for location %s: %s", loc_code, exc)


if __name__ == "__main__":
    main()
