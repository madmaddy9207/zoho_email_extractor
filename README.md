#  Zoho Email Extractor – by Sysdevcode

This is a simple Python tool that extracts **email contacts, sender names, subjects, timestamps, and attachments** from your **Zoho Mail inbox** and saves everything into Excel, CSV, and JSON formats — with analytics included.

We built this because our small startup, **Sysdevcode (Kerala, India)**, got over **1000 inters applications** via Zoho Mail — and we had no auto-reply, no filters, and no plan for follow-ups.  
So, instead of depending on expensive automation tools, we built our own. Now it’s open-source 💚

---

##  Why we created this

After posting an internship poster, we received 1000+ resumes on Zoho Mail.  
But we forgot to set any autoresponders or filters 😅

We looked for tools → Most were paid or didn’t fit.  
We asked AI → It helped, but the logic wasn’t perfect.  
So I decided to code our own tool from scratch.

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
https://github.com/madmaddy9207/zoho_email_extractor.git
cd zoho-email-extractor
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt

```

### 3. Create a Zoho OAuth App

Go to Zoho API Console

Register a new client

Set redirect URI as:

```bash
http://localhost:5000/oauth/callback
```

Save your Client ID and Client Secret

### 4. Set Your Credentials

Linux/macOS:
```bash
export ZOHO_CLIENT_ID="your_client_id"
export ZOHO_CLIENT_SECRET="your_client_secret"
```

Windows PowerShell:

```powershell
$env:ZOHO_CLIENT_ID="your_client_id"
$env:ZOHO_CLIENT_SECRET="your_client_secret"
$env:ZOHO_REDIRECT_URI="http://localhost:5000/oauth/callback"
```

```cmd
set ZOHO_CLIENT_ID="your_client_id"
set ZOHO_CLIENT_SECRET="your_client_secret"
set ZOHO_REDIRECT_URI="http://localhost:5000/oauth/callback"
```

###  Run the Extractor

```bash
python zoho_email_extractor.py
```

It opens your browser to log in via Zoho

Auth flow completes via localhost callback

Data extraction begins in your terminal


### Optional Command-line Flags

--max-messages → Max emails to process (default: 5000)

--no-attachments → Skip downloading attachments

--batch-size → Emails per API request (default: 50)

### Output Directory
All files are saved to zoho_email_extraction/ folder:

| File/Folder            | Description                                  |
| ---------------------- | -------------------------------------------- |
| `contacts_latest.xlsx` | Email list with summary & domain-level stats |
| `contacts_latest.json` | Full metadata with attachments               |
| `contacts_latest.csv`  | Flat email list for bulk tools               |
| `/attachments/`        | Organized by sender email and file type      |


### Security Notes

 - Never commit tokens.json to Git

 - Attachments scanned only by extension (not virus-checked)

 - OAuth tokens are stored securely and auto-refreshed

 - All sensitive info is stored in memory only

### Customization

You can tweak these inside zoho_email_extractor.py:

Increase message cap:

``` python
self.max_messages = 10000
```

Add more attachment types:

```python
self.allowed_extensions.add('.pptx')
```
Slow down rate limit:

```python
self.requests_per_minute = 30
```

### Troubleshooting

| Problem            | Fix                                      |
| ------------------ | ---------------------------------------- |
| `401 Unauthorized` | Delete `tokens.json` and re-authenticate |
| Attachment issues  | Use `--no-attachments` to skip           |
| Slow extraction    | Lower `--batch-size` or increase delay   |
| Missing emails     | Check if correct folder is being used    |

Logs are saved in: zoho_extractor.log


### Contribute

This tool is open-source under the MIT License.
Feel free to fork, contribute, improve, or suggest features.

 Author: Abin P
 Ping us with ideas or bugs
 Made with love by Sysdevcode, Kerala 🇮🇳


### License

MIT License.
This tool is not affiliated with Zoho Corp.

### Compliance Notes

"Not affiliated with Zoho Corp"

"Users must comply with local data privacy laws (GDPR/CCPA)"

```Markdown
WARNING: This tool accesses sensitive email data.  
• Only use on accounts you own  
• Never extract contacts without permission  
• Securely delete downloaded attachments after processing
```
