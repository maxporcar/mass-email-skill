# mass-email — Claude Code Skill

Automated B2B email campaigns with personalization, file attachments, multi-language sending, and delivery logging.

**Think:** Word mail merge × Excel contact list × Outlook/Gmail — but fully automated, with file attachments and zero manual work.

## Skills included

| Skill | Trigger | What it does |
|-------|---------|--------------|
| `mass-email` | "send campaign to my contact list" | Core: personalized send loop with attachments, multi-language, CSV log |
| `mass-email:template` | "extract email template from this DOCX" | Reads a .docx and extracts HTML preserving table colors and formatting |
| `mass-email:setup` | "set up email sending" | Provider wizard: Outlook COM, Gmail SMTP, or generic SMTP + test send |
| `mass-email:audit` | "check bounces", "campaign report" | Parses CSV log + scans inbox for NDR/bounce messages |

## Features

- **Personalization** — `{company}`, `{contact_name}`, any column from your list
- **Multi-language** — send ZH+EN (or any combo) per company in one loop
- **Multi-email per contact** — handles `a@x.com; b@x.com` in one cell
- **Segment templates** — different template per segment column (cosmetics vs food vs care)
- **Providers** — Outlook COM (Windows), Gmail SMTP, generic SMTP
- **Attachments** — PDFs, presentations, any file
- **CSV log** — every send logged with timestamp, status, error
- **DOCX → HTML** — extract colored table boxes from Word templates

## Installation

```bash
# Clone into your Claude skills directory
git clone https://github.com/maxporcar/mass-email-skill ~/.claude/skills/mass-email-skill

# Or install individual skills
cp -r mass-email-skill/mass-email ~/.claude/skills/
cp -r mass-email-skill/mass-email-template ~/.claude/skills/
cp -r mass-email-skill/mass-email-setup ~/.claude/skills/
cp -r mass-email-skill/mass-email-audit ~/.claude/skills/
```

## Quick start

1. **First time?** → Ask Claude: *"set up email sending"* → runs `mass-email:setup`
2. **Have a DOCX template?** → *"extract email template from my template.docx"* → runs `mass-email:template`
3. **Send campaign** → *"send this campaign to my contacts.xlsx"* → runs `mass-email`
4. **After sending** → *"check bounces from last campaign"* → runs `mass-email:audit`

## Contact list format

Excel or CSV:

| company | contact_name | email | language | segment |
|---------|-------------|-------|----------|---------|
| Acme Corp | John Smith | john@acme.com | EN | cosmetics |
| 台灣公司 | 王小明 | wang@tw.com | ZH;EN | food |

- Multiple emails: separate with `;`
- Multiple languages: separate with `;`  
- Add any column — available as `{column_name}` in template body

## Requirements

```bash
pip install pandas openpyxl pywin32  # for Outlook COM
# pywin32 only needed for Outlook — not required for Gmail/SMTP
```

## Providers

- **Outlook COM** — Windows + Outlook Desktop (corporate accounts, supports full signature with logo)
- **Gmail SMTP** — requires App Password (not regular password)
- **Generic SMTP** — any mail server (Office 365, Zoho, self-hosted)

## Tips

- Always test with `--dry-run` first, then `--limit 1` to your own email
- Keep attachments under 10MB total per email
- For ZH+EN campaigns: use 30–45s delay between languages (same recipient)
- Use 5–8s delay between different companies

## Built from real campaigns

This skill was built and battle-tested sending 120+ emails across 60+ Taiwanese B2B companies, handling Traditional Chinese + English bilingual campaigns with PDF attachments and corporate Outlook signatures.

---

MIT License · Made by [@maxporcar](https://github.com/maxporcar)
