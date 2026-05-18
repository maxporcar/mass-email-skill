---
name: mass-email:setup
description: Interactive setup wizard for mass-email campaigns. Installs required Python dependencies, configures the email provider (Outlook COM, Gmail SMTP, or generic SMTP), and sends a test email to verify everything works before launching a real campaign. Use this whenever someone wants to get started with mass-email for the first time, needs to configure their email provider, gets errors running a mass-email script, or says "set up email sending" / "configure Outlook" / "configure Gmail".
---

# mass-email:setup — Email Provider Setup Wizard

Run this before your first campaign. It checks your environment, configures your provider, and sends a test email so you know everything works.

## Step 1: Check Python dependencies

```python
# check_deps.py
import subprocess, sys

REQUIRED = {
    "pandas":    "pip install pandas openpyxl",
    "win32com":  "pip install pywin32",   # Outlook COM only
}

for pkg, install_cmd in REQUIRED.items():
    try:
        __import__(pkg.split(".")[0])
        print(f"  ✅ {pkg}")
    except ImportError:
        print(f"  ❌ {pkg} — install with: {install_cmd}")
```

Run: `python check_deps.py`

## Step 2: Choose your provider

### A) Outlook COM (Windows + Outlook Desktop)

**Best for:** corporate accounts, signature with logo, Windows only.

**Requirements:** Outlook desktop installed and configured with your account.

**Test:**
```python
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = "your@email.com"
mail.Subject = "mass-email test"
mail.Body = "It works!"
mail.Send()
print("Sent via Outlook COM ✅")
```

**Find your signature:**
```python
import os, glob
sig_dir = os.path.expanduser(r"~\AppData\Roaming\Microsoft\Signatures")
print("Signatures found:")
for f in glob.glob(sig_dir + "\\*.htm"):
    print(f"  {f}")
```

**Critical rules (don't skip these):**
- Never call `mail.GetInspector` followed by `mail.Send()` — causes error `-2147024809`
- Never call `mail.Display()` before `Send()` — email ends up in Drafts, not Sent
- Always copy attachments to `C:\Temp\` first — paths with spaces or accents fail silently

---

### B) Gmail SMTP

**Best for:** personal Gmail, G Suite / Google Workspace.

**Requirements:** App Password (not your regular password).

**Get App Password:**
1. myaccount.google.com → Security → 2-Step Verification → App Passwords
2. Create one for "Mail" + "Windows Computer"
3. Copy the 16-character password

**Test:**
```python
import smtplib
from email.mime.text import MIMEText

sender = "you@gmail.com"
app_password = "xxxx xxxx xxxx xxxx"  # 16-char App Password

msg = MIMEText("It works!", "plain", "utf-8")
msg["From"] = sender
msg["To"] = sender
msg["Subject"] = "mass-email test"

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender, app_password)
    server.sendmail(sender, [sender], msg.as_string())
print("Sent via Gmail ✅")
```

**Limits:** 500/day (free) · 2000/day (Workspace)

---

### C) Generic SMTP

**Best for:** any mail server (Office 365, Zoho, self-hosted, etc.).

```python
import smtplib
from email.mime.text import MIMEText

SMTP_HOST = "smtp.yourserver.com"
SMTP_PORT = 587
USER = "you@yourserver.com"
PASSWORD = "yourpassword"

msg = MIMEText("It works!", "plain", "utf-8")
msg["From"] = USER
msg["To"] = USER
msg["Subject"] = "mass-email test"

with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
    server.starttls()
    server.login(USER, PASSWORD)
    server.sendmail(USER, [USER], msg.as_string())
print("Sent via SMTP ✅")
```

## Step 3: Verify test email arrived

Check your inbox. If it didn't arrive:
- **Outlook:** check Outbox (sometimes emails queue) → force sync with `ns.SendAndReceive(False)`
- **Gmail:** check Spam folder · verify App Password is correct
- **SMTP:** verify host/port · try port 465 with SSL instead of 587 with TLS

## Step 4: Ready

Once the test email arrives, you're set. Run `mass-email` to launch your campaign.

## Daily send limits reference

| Provider | Limit |
|----------|-------|
| Outlook (corporate) | ~500/day (varies by server) |
| Gmail free | 500/day |
| Gmail Workspace | 2,000/day |
| SendGrid free | 100/day |
| Mailgun free | 100/day |
| Generic SMTP | depends on your server |

For campaigns > 500 contacts, split across multiple days or use SendGrid/Mailgun.
