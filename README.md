# 📧 mass-email — AI Email Campaign Skill

> **Automated B2B email campaigns with personalization, file attachments, multi-language sending, and delivery logging.**

Works with **Claude Code**, **GitHub Copilot**, **Cursor**, **Gemini CLI**, or any AI agent that supports skills/instructions. Think: **Word mail merge × Excel × Outlook/Gmail** — but fully automated, with attachments, and zero manual work.

---

## ✨ What it does

You give it:
- 📋 A contact list (Excel or CSV)
- 📄 An email template (DOCX or HTML)
- 📎 Attachments (PDFs, presentations...)
- 🌍 Languages (EN, ZH, ES, IT... or multiple per company)

It sends personalized emails to every contact, handles multiple email addresses per company, logs everything to CSV, and tells you what bounced.

---

## 🧩 Skills included

| Skill | Trigger | What it does |
|-------|---------|--------------|
| `mass-email` | *"send campaign to my contact list"* | Core send loop — personalized, multi-language, logged |
| `mass-email:template` | *"extract email template from this DOCX"* | Reads .docx → clean HTML preserving table colors & formatting |
| `mass-email:setup` | *"set up email sending"* | Provider wizard: Outlook COM, Gmail SMTP, or generic SMTP |
| `mass-email:audit` | *"check bounces / campaign report"* | Parse CSV log + scan inbox for NDR/bounce messages |

---

## 📬 Real example (anonymized)

Contact list (`contacts.xlsx`):

| company | contact_name | email | language | segment |
|---------|-------------|-------|----------|---------|
| XXX Cosmetics Co. | Ms. XXX | info@xxx.com.tw; sales@xxx.com.tw | ZH;EN | cosmetics |
| XXX Biotech Ltd. | XXX Team | rd@xxx.com | EN | food |
| XXX International | Mr. XXX | xxx@xxx.com | ZH | care |

Email body with `{placeholders}`:

```html
<p>Dear <strong>{contact_name}</strong>,</p>
<p>After reviewing <strong>{company}</strong>'s product portfolio,
we believe there is an interesting collaboration opportunity...</p>
```

Result → each company gets a personalized email in their language(s), with the right attachments for their segment.

---

## 🎨 Colored proposal boxes (from DOCX)

The skill replicates colored table boxes from Word templates exactly:

```python
def box(bg_hex, border_hex, content_html):
    return (
        f'<table cellpadding="0" cellspacing="0" border="0" width="100%" '
        f'style="border-collapse:collapse;margin:8px 0;background:#{bg_hex};">'
        f'<tr><td style="border-left:4px solid #{border_hex};padding:12px 14px;'
        f'font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">'
        f'{content_html}</td></tr></table>'
    )

# Example — 3 colored proposal boxes
box("F4FBF7", "0A8F55", "<strong>1. Custom solution</strong><br>Tailored to your matrix...")
box("F5F8FB", "294F7A", "<strong>2. Technology</strong><br>Proprietary microencapsulation...")
box("FBF9F2", "7A5A1F", "<strong>3. European Quality</strong><br>FSSC 22000, fast sampling...")
```

---

## 🚀 Quick start

### 1. Install

```bash
# Clone into your AI skills directory
git clone https://github.com/maxporcar/mass-email-skill

# Claude Code
cp -r mass-email-skill/mass-email* ~/.claude/skills/

# Or add to your agent's context/instructions folder
```

### 2. Setup (first time only)
Ask your AI: *"set up email sending"* → it runs `mass-email:setup`, installs deps, configures your provider, sends a test.

### 3. Prepare your template
Have a DOCX? → *"extract email template from my template.docx"* → outputs ready-to-use HTML.

### 4. Send
*"Send this campaign to contacts.xlsx with these PDFs attached"* → AI builds and runs the script.

### 5. Audit
24h later: *"Check bounces from my last campaign"* → report + bounce list.

---

## 📋 Contact list format

Excel or CSV with these columns:

| Column | Required | Notes |
|--------|----------|-------|
| `company` | ✅ | Used in `{company}` placeholder |
| `contact_name` | — | Used in `{contact_name}` |
| `email` | ✅ | Multiple: `a@x.com; b@x.com` |
| `language` | — | `EN`, `ZH`, `ES`... Multiple: `ZH;EN` |
| `segment` | — | Selects template per row |
| `template` | — | Path to DOCX/HTML for this row |
| *any column* | — | Available as `{column_name}` in body |

---

## ⚙️ Email providers

| Provider | Best for | Setup |
|----------|----------|-------|
| **Outlook COM** | Windows + corporate accounts | Outlook Desktop installed |
| **Gmail SMTP** | Personal / Google Workspace | App Password required |
| **Generic SMTP** | Any server (Office 365, Zoho...) | host + port + credentials |

---

## 🤖 Compatible with

This skill works with any AI agent that can read instructions and run Python:

- **Claude Code** — drop folders in `~/.claude/skills/`
- **GitHub Copilot** / **Cursor** — add to `.github/copilot-instructions.md` or rules
- **Gemini CLI** — add to agent instructions
- **Any LLM agent** — paste SKILL.md content into system prompt or context

The Python scripts run standalone — no special framework needed.

---

## 📦 Requirements

```bash
pip install pandas openpyxl   # contact list reading (all providers)
pip install pywin32            # Outlook COM only (Windows)
```

Python 3.8+ · No external API keys needed

---

## ⚡ Delay guidelines

| Situation | Delay |
|-----------|-------|
| Between different companies | 5–8s |
| Between languages (same company) | 30–45s |
| High volume (>200/day) | 15–30s |

---

## ⚠️ Outlook COM gotchas

```python
# ❌ NEVER — causes error -2147024809
mail.GetInspector
mail.Send()

# ❌ NEVER — email ends up in Drafts
mail.Display()
mail.Send()

# ✅ ALWAYS copy attachments to C:\Temp\ first
# Paths with spaces or accents fail silently
```

---

## 📊 Daily send limits

| Provider | Limit |
|----------|-------|
| Outlook (corporate) | ~500/day |
| Gmail free | 500/day |
| Gmail Workspace | 2,000/day |
| Generic SMTP | depends on server |

---

## 🧪 Battle-tested

Built from real B2B prospecting campaigns:
- ✅ 120+ emails sent across 60+ companies in a single run
- ✅ Traditional Chinese + English bilingual campaigns
- ✅ Multiple PDFs per email, multi-email per company
- ✅ Corporate Outlook signature with embedded logo (CID)
- ✅ DOCX template with colored boxes → pixel-perfect HTML email

---

## 📄 License

MIT — free to use, modify, and share.

---

*Made by [@maxporcar](https://github.com/maxporcar) · contributions welcome*
