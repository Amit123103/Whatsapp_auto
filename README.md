# ⚡ Skytecher WhatsApp Outreach Suite (v2.0)

A complete, production-ready Python desktop application with a modern Tkinter UI, **AI-Powered Message Generation** (NVIDIA NIM / OpenAI-compatible API), and automated cold WhatsApp messaging from Excel spreadsheets.

Built specifically for **Skytecher** (IT & Digital Services):
- 🌐 Website: [https://skytecher.com](https://skytecher.com)
- ✉️ Email: `skytechersolutions@gmail.com`
- 📱 Sender WhatsApp Account: `+91 8960061745`

---

## 🎯 Key Capabilities & What's New in v2.0

1. **Intelligent Excel Processing**:
   - Reads `.xlsx` files with columns: `Name`, `Phone`, `Company`, `Service`, `Suggestion`.
   - Cleans and formats phone numbers automatically: adds `+91` for 10-digit Indian numbers, strips spaces/dashes/parentheses, and strips `.0` float artifacts.
   - Detects and skips empty, invalid, and duplicate numbers.

2. **Resilient Non-WhatsApp Number Handling**:
   - If a number does not have WhatsApp or fails to connect, the application catches the error, marks it as `Failed (No WhatsApp / Error)`, logs it in the live console, and **automatically proceeds to the next number** without interrupting the outreach workflow.

3. **🤖 AI-Powered Cold Message Generation (NVIDIA NIM / LLM)**:
   - Integrates with high-performance open models on NVIDIA Build (`build.nvidia.com`):
     - `deepseek-ai/deepseek-v4.1-flash`
     - `openai/gpt-oss-20b`
     - `nvidia/nemotron-3.5-lightning-30b-a3b`
     - `meta/llama-3.3-70b-instruct`
     - Any custom OpenAI-compatible model ID
   - Dynamically crafts a hyper-personalized message incorporating the client's **custom Suggestion**, **Service**, and **Company** alongside Skytecher's full credentials.
   - One-click **"✨ Test AI Generation for Lead #1"** to preview output before starting.
   - If the AI API is disabled or times out, it automatically falls back to the high-converting standard template.

4. **Environment Configuration (`.env`)**:
   - Centralizes owner phone, company details, and AI API keys in [`.env`](file:///c:/Users/amita/myprojects/whatsapp_lead/.env).
   - Easily save or update API keys directly inside the desktop interface.

5. **Anti-Ban Safety Controls**:
   - Configurable **Daily Send Limit** (default: 30 contacts/day).
   - Randomized delay between messages (default: 30–60 seconds).
   - **Simulation / Dry Run Mode** checkbox: test parsing, AI drafting, and report generation without opening browser tabs or sending live messages.

6. **Recipient History & Blocklist Management**:
   - [`sent_history.txt`](file:///c:/Users/amita/myprojects/whatsapp_lead/sent_history.txt): Records every number messaged so no client receives the message twice.
   - [`blocklist.txt`](file:///c:/Users/amita/myprojects/whatsapp_lead/blocklist.txt): Automatically excludes any contact that replies "STOP" (manageable via the in-app **"🛡️ Blocklist Manager"**).

7. **Multi-Threaded UI & Live Progress**:
   - Dispatch runs in a dedicated background worker thread so the UI never freezes.
   - Live color-coded terminal log: `[SUCCESS]`, `[SKIP]`, `[ERROR]`, `[INFO]`.
   - Real-time KPI counters (Total, Sent, Skipped, Failed, Remaining) and animated progress bar.
   - One-click export to **`results.xlsx`**.

---

## 📋 Project Structure

```text
whatsapp_lead/
├── app.py                  # Main complete application (Tkinter UI + Worker + AI Engine)
├── requirements.txt        # Python library dependencies
├── .env                    # Environment settings (Owner phone, company info, AI API key)
├── .env.example            # Environment configuration template
├── sample_leads.xlsx       # Ready-to-test sample lead spreadsheet
├── blocklist.txt           # Blocklisted / STOP phone numbers
├── sent_history.txt        # History log of messaged numbers
├── create_sample_excel.py  # Script to generate sample spreadsheet
├── test_app.py             # Unit tests for cleaner, formatter, and env variables
└── README.md               # User guide and setup instructions
```

---

## 🚀 Quick Setup & Step-by-Step Run Instructions

### Step 1: Install Dependencies
Open PowerShell or your command prompt in this directory:

```bash
pip install -r requirements.txt
```

### Step 2: Configure `.env` (Optional AI Setup)
Open [`.env`](file:///c:/Users/amita/myprojects/whatsapp_lead/.env) to review or add your free NVIDIA API Key:
```ini
SENDER_PHONE=8960061745
COMPANY_NAME=Skytecher
COMPANY_WEBSITE=https://skytecher.com
COMPANY_EMAIL=skytechersolutions@gmail.com

NVIDIA_API_KEY=nvapi-your-key-here
AI_MODEL=deepseek-ai/deepseek-v4.1-flash
ENABLE_AI_PERSONALIZATION=False
```
*(You can also paste the API key directly in the desktop app UI and click "Save to .env".)*

### Step 3: Ensure WhatsApp Web is Logged In
1. Open your default web browser (Chrome, Edge, Brave, etc.).
2. Go to **[https://web.whatsapp.com](https://web.whatsapp.com)**.
3. Log in with your Skytecher WhatsApp number (`8960061745`). Keep the browser open in the background.

### Step 4: Run the Application
```bash
python app.py
```

---

## 🖥️ How to Use the App

1. **Upload Leads or Download Template**:
   - Click **`📥 Download Template`** to save the official 17-column Excel template (`skytecher_leads_template.xlsx`).
   - Click **`📂 Choose Excel File`** to upload your filled leads sheet.
   - Click **`📄 Load Sample Leads`** to test directly with pre-loaded leads.
2. **AI Personalization (Optional)**:
   - Check **"Enable AI Dynamic Personalization"**.
   - Select your model (e.g. `deepseek-ai/deepseek-v4.1-flash` or `openai/gpt-oss-20b`).
   - Click **"✨ Test AI Generation for Lead #1"** to preview how the AI personalizes the message based on that contact's suggestion.
3. **Review / Customize Message Template**:
   - Choose any of the 14 pre-built English & Hinglish formats from the **"Choose Format"** dropdown.
   - Placeholders supported: `{name}`, `{business}`, `{service}`, `{suggestion}`, `{city}`, `{industry}`, `{website}`, `{email}`.
4. **Configure Settings**:
   - Set **Daily Send Limit** (default: 30).
   - Set **Delay Range** (default: 30 to 60 seconds).
   - *(Optional)* Check **`🧪 Simulation / Dry Run`** to test before real live sending.
5. **Start Dispatch**:
   - Click **`▶ START SENDING`**.
   - If a number does not have WhatsApp, the app logs it and automatically moves on to the next contact.
   - Click **`⏹ STOP`** anytime to halt safely.
6. **Export Report**:
   - Click **`📊 Export Results (.xlsx)`** to save the timestamped log for all rows.

---

## 📊 Supported 17-Column Excel Format

Your `.xlsx` file can directly use your complete CRM/lead generation columns:

| # | Column Name | Description & Usage |
| :---: | :--- | :--- |
| 1 | `#` | Row or lead ID |
| 2 | `Company` | Client Business Name (used in `{business}` and `{company}`) |
| 3 | `Industry` | Domain/Sector (used in `{industry}`) |
| 4 | `Lead Type` | Lead Classification (Cold Prospect, Inbound, etc.) |
| 5 | `Country` | Country (`{country}`) |
| 6 | `City` | City/Location (`{city}`) |
| 7 | `Phone` | Mobile/WhatsApp number (cleans floats, +91, dashes) |
| 8 | `Website` | Client's current website (`{client_website}`) |
| 9 | `Website Status` | Performance note (e.g. Slow Loading, Needs Redesign) |
| 10 | `LinkedIn (company)` | Company LinkedIn profile |
| 11 | `Google Rating` | Client's Google Maps rating (`{rating}`) |
| 12 | `Source` | Acquisition channel (Google Maps, Referral, etc.) |
| 13 | `Suggested Services` | Recommended Skytecher service (`{service}`) |
| 14 | `Contact Score` | Internal lead qualification score |
| 15 | `Priority` | Priority level (High, Medium, Urgent) |
| 16 | `Outreach Status` | Pipeline status (New, Pending, etc.) |
| 17 | `Notes` | Specific custom pitch or idea (`{suggestion}`) |

