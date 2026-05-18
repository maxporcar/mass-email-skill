---
name: mass-email:audit
description: Campaign audit combining delivery report and bounce check. Reads the CSV log from a mass-email campaign and scans the inbox for bounce/NDR messages, producing a summary of what was sent, what bounced, and what needs follow-up. Use this after a campaign finishes, when someone asks "how many emails were sent", "did any bounce", "show me the campaign results", "check for bounces", or "any delivery errors".
---

# mass-email:audit — Campaign Report + Bounce Check

Two things in one: parse the campaign CSV log for stats, and scan the inbox for bounce messages (NDRs / undeliverable notices).

## Part 1: Parse the campaign log

Every mass-email campaign produces a CSV log at `outputs/mass_send_log_<timestamp>.csv`.

```python
# audit_campaign.py
import sys, pandas as pd
from datetime import datetime
sys.stdout.reconfigure(encoding="utf-8")

def audit_log(csv_path):
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    print(f"=== Campaign Log: {csv_path} ===\n")

    total = len(df)
    ok = len(df[df["status"] == "OK"])
    errors = len(df[df["status"] == "ERROR"])
    dry = len(df[df["status"] == "DRY"])

    print(f"Total:   {total}")
    print(f"✅ Sent:  {ok}")
    print(f"❌ Error: {errors}")
    if dry: print(f"🔵 Dry:   {dry}")
    print()

    if errors > 0:
        print("--- Errors ---")
        err_df = df[df["status"] == "ERROR"][["company", "recipients", "language", "error"]]
        print(err_df.to_string(index=False))
        print()

    # By language
    if "language" in df.columns:
        print("--- By language ---")
        print(df.groupby("language")["status"].value_counts().to_string())
        print()

    # Timing
    if "timestamp" in df.columns:
        df["ts"] = pd.to_datetime(df["timestamp"], errors="coerce")
        valid = df[df["ts"].notna()]
        if len(valid) > 1:
            duration = (valid["ts"].max() - valid["ts"].min()).total_seconds() / 60
            print(f"Duration: {duration:.1f} minutes")
            print(f"First:    {valid['ts'].min()}")
            print(f"Last:     {valid['ts'].max()}")

if __name__ == "__main__":
    audit_log(sys.argv[1] if len(sys.argv) > 1 else "outputs/mass_send_log_latest.csv")
```

Run: `python audit_campaign.py outputs/mass_send_log_20240515_142300.csv`

## Part 2: Check for bounces in inbox

Bounce messages (NDRs) arrive in your inbox hours or days after sending. Check these keywords:

```python
# check_bounces.py — Outlook COM version
import win32com.client, sys
sys.stdout.reconfigure(encoding="utf-8")

BOUNCE_KEYWORDS = [
    "undeliverable", "undelivered", "delivery failed",
    "delivery status", "mail delivery", "non-delivery",
    "bounced", "mailer-daemon", "postmaster"
]

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)  # olFolderInbox

bounces = []
for item in inbox.Items:
    try:
        subject = (item.Subject or "").lower()
        sender = (item.SenderEmailAddress or "").lower()
        if any(k in subject for k in BOUNCE_KEYWORDS) or "mailer-daemon" in sender:
            bounces.append({
                "received": item.ReceivedTime,
                "from": item.SenderName,
                "subject": item.Subject[:80],
            })
    except:
        pass

if bounces:
    print(f"⚠️  {len(bounces)} bounces found:\n")
    for b in bounces:
        print(f"  {b['received']} | {b['from'][:30]} | {b['subject']}")
else:
    print("✅ No bounces found in inbox.")
```

For Gmail/SMTP, bounces arrive as regular emails from `mailer-daemon@...` or `postmaster@...`. Check your inbox manually or use IMAP:

```python
# check_bounces_imap.py — Gmail / generic IMAP version
import imaplib, email
from email.header import decode_header

def check_bounces_gmail(gmail_user, app_password):
    with imaplib.IMAP4_SSL("imap.gmail.com") as m:
        m.login(gmail_user, app_password)
        m.select("INBOX")
        _, ids = m.search(None, 'FROM "mailer-daemon"')
        ids2 = m.search(None, 'SUBJECT "undeliverable"')[1]
        all_ids = set(ids[0].split() + ids2[0].split())
        print(f"Bounces found: {len(all_ids)}")
        for mid in list(all_ids)[:10]:
            _, data = m.fetch(mid, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])
            print(f"  From: {msg['From']} | Subject: {msg['Subject'][:60]}")
```

## Summary output format

After running both parts, summarize like this:

```
=== Campaign Audit ===
Sent:       120 emails (60 companies × ZH+EN)
Errors:     2 (company X — connection timeout)
Bounces:    3 (bad addresses: a@b.com, c@d.com, e@f.com)
Success:    97% (115/120 delivered)

Action needed:
→ Fix 2 script errors and resend
→ Remove 3 bounced addresses from contact list
```

## When to run

- Immediately after campaign: check script errors in log
- 24–48 hours after: check bounces (some servers are slow)
- Before next campaign: clean bounced addresses from contact list
