"""
Mailer module for the NIBSS ETL pipeline.

Sends two emails via SMTP (Microsoft 365 / Outlook) after a successful run:

1. Team email — the generated workbook attached
2. MD email   — executive summary of headline KPIs (workbook attached too)

Configuration is read from environment variables:

    EMAIL_USER       your Outlook / Microsoft 365 address (sender)
    EMAIL_PASSWORD   app password for that mailbox
    TEAM_EMAIL       recipient for the workbook email (comma-separated allowed)
    MD_EMAIL         recipient for the executive summary email

If any variable is missing the emails are skipped with a warning —
the pipeline never crashes over mail.  Uses ``smtplib`` + ``email``
only; no third-party dependencies.
"""

import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
import pandas as pd

SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587

REQUIRED_ENV = ("EMAIL_USER", "EMAIL_PASSWORD", "TEAM_EMAIL")


def _executive_summary(data: dict) -> dict:
    """Headline KPIs for the MD email, from both sources combined."""
    parts = [
        df for key in ("easy_pay", "direct_debit")
        if (df := data.get(key)) is not None and not df.empty
        and "status" in df.columns and "amount" in df.columns
    ]
    if not parts:
        return {}
    combined = pd.concat(parts, ignore_index=True)
    total = len(combined)
    ok = combined[combined["status"] == "successful"]
    total_val = float(combined["amount"].fillna(0).sum())
    ok_val = float(ok["amount"].fillna(0).sum())
    return {
        "Total Transactions": f"{total:,}",
        "Total Approved": f"{len(ok):,}",
        "Total Failed": f"{total - len(ok):,}",
        "Success Rate": f"{len(ok) / total * 100:.1f}%" if total else "n/a",
        "Total Value": f"₦{total_val:,.2f}",
        "Approved Value": f"₦{ok_val:,.2f}",
        "Failed Value": f"₦{total_val - ok_val:,.2f}",
    }


def _render_table(summary: dict) -> str:
    """KPI dict → HTML table rows."""
    return "".join(
        f"<tr><td style='padding:4px 14px'>{k}</td>"
        f"<td align='right' style='padding:4px 14px'>{v}</td></tr>"
        for k, v in summary.items()
    )


def _compose(title: str, paragraphs: list[str], summary: dict,
             footer: str | None = None) -> tuple[str, str]:
    """Return ``(text, html)`` mail bodies: heading, paragraphs, KPI table, optional footer."""
    text = "\n\n".join(paragraphs)
    text += "\n\n" + "\n".join(f"{k}: {v}" for k, v in summary.items())
    if footer:
        text += f"\n\n{footer}"
    html = (
        "<html><body style='font-family:Segoe UI, Arial, sans-serif;font-size:14px'>"
        + (f"<h2 style='color:#1F4E79'>{title}</h2>" if title else "")
        + "".join(f"<p>{p}</p>" for p in paragraphs)
        + "<table border='1' cellspacing='0' style='border-collapse:collapse'>"
        f"{_render_table(summary)}</table>"
    )
    if footer:
        html += f"<p style='color:#808080;font-size:11px'>{footer}</p>"
    return text, html + "</body></html>"


def _build_message(subject: str, recipients: list[str], html_body: str,
                   text_body: str, attach_path: str | None = None) -> EmailMessage:
    """Wrap HTML + text bodies into a multipart message, optionally with the workbook attached."""
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("NIBSS ETL", os.environ["EMAIL_USER"]))
    msg["To"] = ", ".join(recipients)
    msg.set_content(text_body)
    msg.add_alternative(html_body, subtype="html")

    if attach_path:
        with open(attach_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=Path(attach_path).name,
            )
    return msg


def send_emails(data: dict, output_path: str, date_folder: str) -> bool:
    """
    Send the team (workbook) and MD (summary) emails.

    Returns ``True`` when both were sent, ``False`` when env vars were
    missing or SMTP failed (failure is printed, never raised).
    """
    missing = [name for name in REQUIRED_ENV if not os.environ.get(name)]
    if missing:
        print(f"Emails skipped — missing env vars: {', '.join(missing)}")
        return False

    summary = _executive_summary(data)
    if not summary:
        print("Emails skipped — no data to summarise.")
        return False

    date_label = datetime.strptime(date_folder, "%d_%m_%Y").strftime("%d %B %Y")
    timestamp = datetime.now().strftime("%d %B %Y, %H:%M")

    team_text, team_html = _compose(
        "",
        [
            "Hello Team,",
            f"Please find attached the NIBSS EasyPay (Outward) reconciliation "
            f"report for {date_label}.",
            "Should any customer call regarding a failed transaction, kindly use "
            "the attached report to verify the transaction status before responding.",
        ],
        summary,
    )
    team_msg = _build_message(
        f"{date_label} NIBSS EasyPay (Outward) Report",
        [e.strip() for e in os.environ["TEAM_EMAIL"].split(",") if e.strip()],
        team_html, team_text,
        attach_path=output_path,
    )

    md_email = os.environ.get("MD_EMAIL", "").strip()
    md_msg = None
    if md_email:
        md_text, md_html = _compose(
            "",
            [
            "Dear Sir,",
            f"Please find below the executive summary of NIBSS EasyPay outward "
            f"transactions for {date_label} for the bank's kind review.",
                "The full report is attached for further drill-down should you need "
                "transaction-level details.",
            ],
            summary,
            footer=f"Generated automatically by the NIBSS EasyPay ETL pipeline at {timestamp}.",
        )
        md_msg = _build_message(
            f"{date_label} Executive Summary — NIBSS EasyPay (Outward) Report",
            [md_email],
            md_html, md_text,
            attach_path=output_path,
        )

    user = os.environ["EMAIL_USER"]
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.login(user, os.environ["EMAIL_PASSWORD"])
            server.send_message(team_msg)
            if md_msg:
                server.send_message(md_msg)
    except Exception as exc:
        print(f"Emails failed: {exc}")
        return False

    print(f"Emails sent — team: {team_msg['To']}")
    if md_msg:
        print(f"Emails sent — MD: {md_msg['To']}")
    else:
        print("Emails sent — MD skipped (MD_EMAIL not set)")
    return True
