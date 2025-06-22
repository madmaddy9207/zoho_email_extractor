Here's your `README.md` ready to copy-paste directly into your GitHub repo for the **Zoho Email Extractor** project:

---

````markdown
# 📥 Zoho Email Extractor – by Sysdevcode

A simple Python tool to extract emails, sender names, subjects, and attachments from your **Zoho Mail inbox** — exports all the data to Excel, JSON, and CSV.

We built this tool after getting **1000+ internship applications** and realizing we had no way to reply or manage them. No auto-reply, no filters, no tools — so we built our own.  
Now it's open-source. MIT licensed. No fancy GUI. Just works.

---

## 🔧 Features

- ✅ OAuth2 login with Zoho
- 📩 Extract sender email, name, subject, timestamp
- 📎 Download attachments (PDF, DOCX, XLSX, etc.)
- 📊 Export to:
  - Excel with stats & charts
  - JSON with full metadata
  - CSV for bulk email tools
- 📈 Domain analytics + sender frequency ranking
- 🕒 Rate-limited (safe for Zoho API)
- 💾 All attachments saved in `/attachments/sender_email/`

---

## 🛠️ Setup

### 1. Clone the repo

```bash
git clone https://github.com/YOUR-USERNAME/zoho-email-extractor.git
cd zoho-email-extractor
````

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Get Zoho OAuth2 credentials

* Go to: [Zoho Developer Console](https://api-console.zoho.in/)
* Create a client
* Set redirect URI to: `http://localhost:5000/oauth/callback`
* Copy your Client ID and Client Secret

### 4. Set environment variables

```bash
export ZOHO_CLIENT_ID="your_client_id"
export ZOHO_CLIENT_SECRET="your_client_secret"
```

You can also add these to a `.env` file.

---

## ▶️ Run the script

```bash
python zoho_email_extractor.py
```

* Opens browser to authenticate
* Starts extracting emails from your inbox
* Saves output to `zoho_email_extraction/` folder

---

## 📂 Output Files

| File                   | Description                                             |
| ---------------------- | ------------------------------------------------------- |
| `contacts_latest.xlsx` | Main Excel file with contacts, stats & domain analytics |
| `contacts_latest.json` | Full metadata including timestamps, attachments         |
| `contacts_latest.csv`  | Clean email list for importing into mail tools          |
| `/attachments/`        | All saved files from emails, organized by sender        |

---

## 🧠 Why we built this

We’re a small startup from Kerala called **Sysdevcode**.
After our first internship drive, we were flooded with emails and no automation. We searched tools, tried GPT, nothing fit our exact need — so we built this from scratch.

Now it’s open-source so others like us can benefit.
Not perfect, but it’s clean and works well.

---

## 🙌 Contribute

* Found a bug? Open an issue.
* Want to improve something? PRs welcome.
* Got a suggestion? Ping us anytime.

> This project is built by @AbinP from **Sysdevcode** – connect on [LinkedIn](https://www.linkedin.com/in/abinp-/)

---

## 📜 License

MIT License.
Not affiliated with Zoho Corp.

---

Made with 💻 in Kerala 🇮🇳

```

---

Let me know if you want:
- A `badge section` (stars, license, etc.)
- Screenshots or a GIF demo for the README
- Deployment as a PyPI or Docker version
```
