#  Zoho Email Extractor – by Sysdevcode

This is a simple but powerful Python tool that extracts **email contacts, sender names, subjects, timestamps, and attachments** from your **Zoho Mail inbox** and saves everything into Excel, CSV, and JSON formats — with analytics included.

We built this because our small startup, **Sysdevcode (Kerala, India)**, got over **1000 internship applications** via Zoho Mail — and we had no auto-reply, no filters, and no plan for follow-ups.  
So, instead of depending on expensive automation tools, we built our own. Now it’s open-source 💚

---

##  Why we created this

After posting an internship poster, we received 1000+ resumes on Zoho Mail.  
But we forgot to set any autoresponders or filters 😅

We looked for tools → Most were paid or didn’t fit.  
We asked AI → It helped, but the logic wasn’t perfect.  
So we decided to code our own tool from scratch.

Now it's working perfectly — and open to anyone who wants to use, fork, or contribute.

---

##  Features

-  Zoho Mail OAuth2 Authentication (secure)
-  Extract email addresses, sender names, subjects, and timestamps
-  Download attachments (PDF, DOCX, XLSX, etc.)
-  Export reports in:
  - Excel with stats and domain analytics
  - JSON with full metadata
  - CSV for email marketing tools
-  Domain-level grouping + sender frequency stats
-  Rate-limited (safe for Zoho API)
-  Save attachments by sender in `/attachments/`

---

##  Technologies Used

- **Python 3.7+**
- `requests`, `pandas`, `openpyxl`, `logging`
- Zoho Mail API & OAuth2
- Local HTTP server for secure OAuth callback
- `.env` or environment variables for credential handling

---

##  Setup

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR-USERNAME/zoho-email-extractor.git
cd zoho-email-extractor
