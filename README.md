# FUR CPR Automation System
**Dragon Studios** for FUR Store (furstores.myshopify.com)

## What it does
Upload a PostEx CPR PDF → auto-matches Shopify orders → marks delivered orders PAID → verifies returns are canceled → downloads Excel report.

---

## Option A: Deploy to Railway (FREE hosting — Recommended)

### Step 1: Upload to GitHub
1. Create account at github.com (free)
2. New repository → name it `fur-cpr-app`
3. Upload all files from this folder

### Step 2: Deploy to Railway
1. Go to railway.app → Sign up with GitHub (free)
2. New Project → Deploy from GitHub repo → select `fur-cpr-app`
3. Railway auto-detects Python and deploys
4. Your app will be live at: `https://fur-cpr-app.up.railway.app`
5. Share this URL with your staff — done!

**Cost: FREE** (500 hours/month free tier — more than enough)

---

## Option B: Run Locally on Any Computer

### Requirements
- Python 3.9+

### Setup (one time per computer)
```bash
pip install -r requirements.txt
python app.py
```
Open browser: http://localhost:5000

---

## How to get Shopify API Token

1. Shopify Admin → Settings → Apps and sales channels
2. Develop apps → Create an app (name: "CPR Automation")
3. Configure Admin API scopes → enable:
   - `write_orders`
   - `read_orders`
4. Install app → Admin API access token → Copy it
5. Paste in the web app

---

## Files
- `app.py` — Flask backend (PDF parsing, Shopify API, Excel export)
- `templates/index.html` — Web UI
- `requirements.txt` — Python packages
- `Procfile` — Railway/Render deployment config

---

## Staff Usage
1. Open the URL in any browser
2. Paste the Shopify API token
3. Upload CPR PDF from PostEx
4. Click "Run Automation"
5. Download Excel report when done

**No installation needed for staff — just a browser!**
