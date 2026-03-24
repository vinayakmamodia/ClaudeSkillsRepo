"""
Daily summary email report — runs via GitHub Actions at 5 PM UTC every day.
Required environment variables (set as GitHub Secrets):
  SMTP_HOST      e.g. smtp.gmail.com
  SMTP_PORT      e.g. 587
  SMTP_USER      sender email address
  SMTP_PASSWORD  sender email password / app password
  RECIPIENT_EMAIL  destination inbox
"""

import os
import smtplib
import csv
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def load_boardline_summary(csv_path="Boardline_Numbers.csv"):
    """Return a quick summary of the Boardline Numbers CSV."""
    rows = []
    try:
        with open(csv_path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            rows = list(reader)
    except FileNotFoundError:
        return "<p>Boardline_Numbers.csv not found in repository.</p>", 0

    total = len(rows)
    # Build an HTML table from the first 10 rows as a preview
    if not rows:
        return "<p>No data found in Boardline_Numbers.csv.</p>", 0

    headers = list(rows[0].keys())
    header_html = "".join(f"<th style='padding:6px 12px;background:#4472C4;color:#fff;text-align:left'>{h}</th>" for h in headers)
    row_html = ""
    for i, row in enumerate(rows[:10]):
        bg = "#f2f6fc" if i % 2 == 0 else "#ffffff"
        cells = "".join(f"<td style='padding:5px 12px;border-bottom:1px solid #ddd'>{row.get(h, '')}</td>" for h in headers)
        row_html += f"<tr style='background:{bg}'>{cells}</tr>"

    table_html = f"""
    <table style='border-collapse:collapse;width:100%;font-size:13px;font-family:Arial,sans-serif'>
      <thead><tr>{header_html}</tr></thead>
      <tbody>{row_html}</tbody>
    </table>
    {"<p style='color:#666;font-size:12px'>Showing first 10 of " + str(total) + " records.</p>" if total > 10 else ""}
    """
    return table_html, total


def build_html_email(table_html, total_records):
    now = datetime.now(timezone.utc).strftime("%A, %d %B %Y — %I:%M %p UTC")
    return f"""
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;color:#333;max-width:800px;margin:auto;padding:20px">
  <div style="background:#4472C4;color:#fff;padding:18px 24px;border-radius:6px 6px 0 0">
    <h2 style="margin:0">Daily Summary Report</h2>
    <p style="margin:4px 0 0;font-size:13px;opacity:.85">{now}</p>
  </div>
  <div style="border:1px solid #ddd;border-top:none;padding:20px;border-radius:0 0 6px 6px">
    <h3 style="margin-top:0">Boardline Numbers Overview</h3>
    <p><strong>Total records:</strong> {total_records}</p>
    {table_html}
    <hr style="margin:24px 0;border:none;border-top:1px solid #eee">
    <p style="font-size:12px;color:#999">
      This is an automated daily report generated from the
      <a href="https://github.com/vinayakmamodia/claudeskillsrepo">ClaudeSkillsRepo</a>.
      Delivered every day at 5:00 PM UTC.
    </p>
  </div>
</body>
</html>
"""


def send_email(subject, html_body, recipient, smtp_host, smtp_port, smtp_user, smtp_password):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = recipient
    msg.attach(MIMEText(html_body, "html"))

    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
        server.ehlo()
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, recipient, msg.as_string())
        print(f"Email sent successfully to {recipient}")


if __name__ == "__main__":
    smtp_host = os.environ["SMTP_HOST"]
    smtp_port = os.environ.get("SMTP_PORT", "587")
    smtp_user = os.environ["SMTP_USER"]
    smtp_password = os.environ["SMTP_PASSWORD"]
    recipient = os.environ["RECIPIENT_EMAIL"]

    table_html, total_records = load_boardline_summary()
    html_body = build_html_email(table_html, total_records)

    today = datetime.now(timezone.utc).strftime("%d %b %Y")
    subject = f"Daily Summary Report — {today}"

    send_email(subject, html_body, recipient, smtp_host, smtp_port, smtp_user, smtp_password)
