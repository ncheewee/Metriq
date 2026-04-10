# MetrIQ MVP v1 — Setup Guide

## Files
- `index.html` — The complete PWA app (single file, deploy anywhere)
- `Code.gs` — Google Apps Script backend
- `SETUP.md` — This guide

## Step 1: Google Apps Script Backend

1. Go to [script.google.com](https://script.google.com) with your new Google account
2. Click **New Project**
3. Delete the default code, paste the contents of `Code.gs`
4. Click **Save** (name it "MetrIQ Backend")
5. Click **Run → setupSheets** (first time only — creates the 4 sheets + default config)
6. Go to **Deploy → New Deployment**
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
7. Click **Deploy** → copy the **Web App URL**

## Step 2: Get Your API Key

1. Open the deployed Google Sheet (you'll find it in your Drive as "MetrIQ Backend")
2. Go to the **Config** tab
3. Find the `api_key` row — copy its value (it's auto-generated)

## Step 3: Configure the App

1. Open `index.html` in a browser (or host it on GitHub Pages)
2. As **Admin**, tap the Admin role card
3. You need to add the first Admin user directly in the Google Sheet:
   - Go to the **Users** sheet
   - Add a row: `U_001 | Your Name | Admin | [see below] | | your@email.com | Y | [today]`
   - For `pin_hash`: in Apps Script console, run `Logger.log(hashPin('123456'))` to get the hash for PIN `123456`
4. Go to **Settings → Backend** and paste:
   - Web App URL
   - API Key
   - Site Name (e.g. "Block 5A")
5. Save — then go to **Settings → Users** to add more users via the app UI

## Step 4: Add Your First Admin PIN (Shortcut)

Run this in Apps Script to add an admin user directly:
```javascript
function addFirstAdmin() {
  addUser({ name: 'Admin', role: 'Admin', pin: '123456', email: 'your@email.com', apiKey: 'YOUR_API_KEY' });
}
```

## EmailJS Setup

1. Go to [emailjs.com](https://emailjs.com) → Sign up (free)
2. Add an Email Service (Gmail works)
3. Create a template with these variables:
   - `{{to_email}}`, `{{subject}}`, `{{message}}`, `{{site_name}}`, `{{month}}`
4. Copy: Service ID, Template ID, Public Key → paste into Admin → Settings → EmailJS

## Telegram Setup

1. Message [@BotFather](https://t.me/botfather) → `/newbot` → get your Bot Token
2. Add the bot to your group chat
3. Get the group Chat ID from [@userinfobot](https://t.me/userinfobot)
4. Paste both into Admin → Settings → Telegram

## Hosting

Option A — GitHub Pages (free):
- Push `index.html` to a repo → Settings → Pages → Deploy from main branch

Option B — Direct:
- Open `index.html` directly in any browser — works offline too (PWA)

## Role Summary

| Role | Access |
|---|---|
| Inspector | Capture readings, view own meter history |
| Manager | Overview, analytics, send reports, Save As |
| Admin | Everything + manage meters & users |

## Limits

- Max **100 meters** (enforced in app)
- **1 reading per meter per calendar month**
- EmailJS free tier: **200 emails/month** (shown in app)
- Google Drive free: **15 GB** (~3+ years at 100 photos/month × 40KB)
