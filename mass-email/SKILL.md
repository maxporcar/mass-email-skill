---
name: mass-email
description: Automated B2B email campaigns with personalization, file attachments, multi-language sending, and delivery logging. Works like Word mail merge + Excel contact list + Outlook/Gmail — but fully automated with attachments. Use this skill whenever the user wants to send emails to a list of contacts, run a prospecting campaign, do outreach to multiple companies, send personalized emails with PDFs or attachments, or automate any kind of bulk email sending. Also triggers for "envío masivo", "campaña email", "mass mailing", "outreach campaign", "send to my contact list".
---

# mass-email — Automated B2B Email Campaigns

Send personalized emails to a contact list with attachments, multi-language support, and automatic logging. Think: Word mail merge × Excel × Outlook/Gmail, but with file attachments and zero manual work.

## When to use different templates

**Same template for everyone** → one template, loop through list.

**Offers differ significantly by segment** (e.g., cosmetics vs. food vs. home care) → use a `template` column in the contact list pointing to different template files per segment. The script picks the right one per row. Never send a cosmetics pitch to a food company.

## Contact list format

Excel or CSV with these columns (add any custom fields you need):

| Column | Required | Description |
|--------|----------|-------------|
| `company` | ✅ | Company name |
| `contact_name` | optional | Person's name for personalization |
| `email` | ✅ | One or more emails separated by `;` |
| `language` | optional | `EN`, `ZH`, `ES`, etc. — if blank, send all languages |
| `segment` | optional | Used to select template (e.g., `cosmetics`, `food`) |
| `template` | optional | Path to DOCX/HTML template file for this row |
| any other | optional | Available as `{column_name}` in template body |

## Email body personalization

Use `{placeholder}` syntax in the email body HTML. At send time, each placeholder is replaced with the value from that row's column.

```html
<p>Dear <strong>{contact_name}</strong>,</p>
<p>After reviewing {company}'s product portfolio...</p>
```

If a placeholder is missing for a row, it's left blank (no crash).

## Step-by-step workflow

### 1. Understand the campaign
Ask the user:
- What's the contact list? (Excel/CSV path, or paste directly)
- What's the email body? (existing DOCX template, or write from scratch)
- What language(s) per company? (EN only, ZH+EN, etc.)
- Any file attachments? (PDFs, presentations)
- Which email provider? (Outlook on Windows, Gmail, other SMTP)

If the user has a DOCX template → invoke `mass-email:template` to extract the HTML body.
If the user needs provider setup → invoke `mass-email:setup` first.

### 2. Build the script

Generate a Python script at `scripts/send_mass_<campaign_name>.py` using the pattern below. Keep it simple: one file, runs from command line.

**Required arguments:**
```
--contacts   Path to Excel/CSV contact list
--limit N    Test with first N companies (default: all)
--skip N     Resume from company N (default: 0)
--dry-run    Log only, don't actually send
```

**Core sending logic:**
```python
# -*- coding: utf-8 -*-
import sys, os, re, csv, time, datetime, argparse
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

PR_CID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

def parse_emails(cell):
    """Extract all valid emails from a cell (handles ; separators)."""
    return list(dict.fromkeys(
        e.strip().lower()
        for e in re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", str(cell or ""))
    ))

def personalize(html_body, row):
    """Replace {placeholders} with row values. Missing = empty string."""
    for col, val in row.items():
        html_body = html_body.replace(f"{{{col}}}", str(val) if pd.notna(val) else "")
    return html_body

def send_outlook(mail_item, to_list, subject, html_body, attachments, sig_img=None, sig_cid=None):
    mail_item.To = "; ".join(to_list)
    mail_item.Subject = subject
    mail_item.HTMLBody = html_body
    if sig_img and sig_cid:
        att = mail_item.Attachments.Add(sig_img, 1, 1, sig_cid)
        att.PropertyAccessor.SetProperty(PR_CID, sig_cid)
    for path in attachments:
        mail_item.Attachments.Add(path)
    mail_item.Send()
```

**Signature embedding (Outlook COM):**
```python
import re

def load_outlook_signature(htm_path, img_path, cid="sig_img"):
    """Read Outlook .htm signature and replace local image src with CID."""
    with open(htm_path, "r", encoding="windows-1252") as f:
        html = f.read()
    m = re.search(r"<body[^>]*>(.*?)</body>", html, re.DOTALL | re.IGNORECASE)
    body = m.group(1) if m else html
    # Replace any local image src with the CID reference
    body = re.sub(r'src="[^"]*\.png"', f'src="cid:{cid}"', body)
    return body
# Signature image lives at: C:\Users\<user>\AppData\Roaming\Microsoft\Signatures\<name>_archivos\image001.png
# IMPORTANT: always copy attachments to C:\Temp\ first — paths with spaces/accents fail silently in Outlook COM
```

**Colored proposal box (from DOCX):**
```python
def box(bg, border_color, content_html):
    return (
        f'<table cellpadding="0" cellspacing="0" border="0" width="100%" '
        f'style="border-collapse:collapse;margin:8px 0;background:{bg};">'
        f'<tr><td style="border-left:4px solid {border_color};padding:12px 14px;'
        f'font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">'
        f'{content_html}</td></tr></table>'
    )
# Use mass-email:template to extract exact colors from your DOCX
```

**Delay guidelines:**
- Between different companies: **5–8s** (different recipients = low spam risk)
- Between languages for same company: **30–45s** (same recipient = avoid duplicate detection)
- Never use GetInspector + Send() together in Outlook COM — causes error -2147024809

**CSV log (auto-generated):**
```python
# Log every send attempt — success or error
writer.writerow(["timestamp", "company", "contact", "recipients", "language", "subject", "status", "error"])
```

### 3. Test before mass send

Always run with `--dry-run` first to preview the recipient list, then `--limit 1` to test one real send to your own email, then `--limit 5` before full campaign.

### 4. Execute and monitor

Run in background for large lists. Check progress via the CSV log. After completion, run `mass-email:audit` to check for bounces and get a summary.

## Providers

### Outlook COM (Windows corporate accounts)
```python
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
# NEVER call mail.GetInspector then mail.Send() — it breaks
# NEVER call mail.Display() before Send() — item ends up in Drafts
# ALWAYS copy attachments to C:\Temp\ first
```
Requires: `pip install pywin32` · Windows only · Outlook desktop installed

### Gmail SMTP
```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender_email, app_password)  # App Password, not account password
    server.sendmail(sender_email, recipients, msg.as_string())
```
Requires: Gmail App Password (myaccount.google.com → Security → App Passwords)

### Generic SMTP
```python
with smtplib.SMTP("smtp.yourserver.com", 587) as server:
    server.starttls()
    server.login(user, password)
    server.sendmail(sender, recipients, msg.as_string())
```

## Limits to keep in mind
- Outlook corporate: ~500 emails/day typical limit
- Gmail free: 500/day · Gmail Workspace: 2000/day
- Keep attachments under 10MB total per email
- If total PDFs > 8MB, split by language (ZH PDFs for ZH email, EN PDFs for EN email)
