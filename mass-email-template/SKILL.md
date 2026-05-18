---
name: mass-email:template
description: Extract the HTML body from a DOCX email template, preserving table colors, borders, and formatting exactly as designed. Use this whenever someone has a Word document (.docx) they want to use as an email template, needs to convert a DOCX to HTML for sending, wants to replicate the formatting from a Word document in an email campaign, or says "use this Word file as my email template".
---

# mass-email:template — Extract HTML from DOCX

Reads a `.docx` file and outputs a clean HTML body you can drop directly into a mass-email campaign — preserving the colored table boxes, borders, fonts, and layout from the original design.

## How it works

1. Unzip the DOCX (it's a ZIP archive with XML inside)
2. Parse `word/document.xml` to extract paragraphs and tables
3. For each table cell: read the background color (`w:shd/@w:fill`) and left border color (`w:tcBorders/w:left/@w:color`)
4. Convert to inline-CSS HTML table boxes (email-safe format)
5. Output the final HTML body + a `box()` helper function

## Run it

```python
# extract_template.py
import sys, zipfile, re
from xml.etree import ElementTree as ET
sys.stdout.reconfigure(encoding="utf-8")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def extract_template(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    # Extract table colors
    print("=== TABLE BOXES DETECTED ===")
    for tbl in root.iter(f"{W}tbl"):
        for ri, row in enumerate(tbl.iter(f"{W}tr")):
            for cell in row.iter(f"{W}tc"):
                shd = cell.find(f".//{W}shd")
                bg = shd.get(f"{W}fill") if shd is not None else None
                tcb = cell.find(f".//{W}tcBorders")
                left = None
                if tcb is not None:
                    lb = tcb.find(f"{W}left")
                    if lb is not None:
                        left = lb.get(f"{W}color")
                text = "".join(t.text for t in cell.iter(f"{W}t") if t.text)[:60]
                if text.strip():
                    print(f"  bg:#{bg}  border-left:#{left}  | {text}")

    # Extract full text
    texts = [t.text for t in root.iter(f"{W}t") if t.text and t.text.strip()]
    print("\n=== BODY TEXT (first 2000 chars) ===")
    print(" ".join(texts)[:2000])

if __name__ == "__main__":
    extract_template(sys.argv[1])
```

## Output: box() helper

Once you have the colors, use this function to recreate each colored box in HTML:

```python
def box(bg_hex, border_hex, content_html):
    """Email-safe colored table box matching DOCX design."""
    return (
        f'<table cellpadding="0" cellspacing="0" border="0" width="100%" '
        f'style="border-collapse:collapse;margin:8px 0;background:#{bg_hex};">'
        f'<tr><td style="border-left:4px solid #{border_hex};padding:12px 14px;'
        f'font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000;">'
        f'{content_html}</td></tr></table>'
    )
```

## Example output

For a DOCX with 3 colored boxes, you'd get something like:

```
bg:#F4FBF7  border-left:#0A8F55  | 1. Custom Fragrances — tailor-made scents...
bg:#F5F8FB  border-left:#294F7A  | 2. Microencapsulation — long-lasting...
bg:#FBF9F2  border-left:#7A5A1F  | 3. European Quality & Speed — IFRA certified...
```

Then use `box("F4FBF7", "0A8F55", "<p>...</p>")` to recreate each one in your email.

## Notes

- Works with any DOCX: Canva exports, Word files, PowerPoint-to-DOCX
- Images referenced locally (e.g., signature images) must be embedded as CID attachments — see `mass-email` skill for how
- If a table has no colored cells, the DOCX probably uses plain paragraphs — extract as `<p>` elements instead
