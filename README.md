# FitFocus Analytics Dashboard

Vite + React dashboard, auto-synced from a private Excel file in Google Drive via an
Apps Script backend (no separate server — Apps Script IS the backend). Login-gated with
email + password (hashed server-side).

## Before you deploy — 2 things to fill in

### 1. `AppsScript.gs` → Google Apps Script (the backend)

1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Delete the default code, paste in the contents of `AppsScript.gs` from this repo.
3. In your Google Drive, find your `.xlsx` file → copy the file ID from its share URL
   (the long string between `/d/` and `/view`).
4. In the pasted code, find `setup()` and replace `PASTE_YOUR_DRIVE_FILE_ID_HERE` with
   that file ID. (Email + password are already filled in.)
5. Run the `setup` function once (function dropdown → `setup` → ▶ Run). Approve the
   permission prompt — it's your own script reading your own Drive.
6. **Deploy → New deployment → Web app** → Execute as **Me** → Who has access **Anyone**
   → Deploy. Copy the `/exec` URL it gives you.

### 2. `src/FitFocusDashboard.jsx` → paste that URL in

Near the top of the file:
```js
const APPS_SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
```
Replace with the URL from step 6 above.

## Deploy to Vercel

```bash
npm install
npm run dev      # test locally first at localhost:5173
```

Then push this folder to a GitHub repo and import it in Vercel — it auto-detects Vite,
no extra config needed. Or deploy directly from your machine:

```bash
npm i -g vercel
vercel
```

## Changing the login password later

Edit the password inside `setup()` in the Apps Script editor and re-run it. No redeploy
needed — Script Properties update immediately.

## Local dev note

`npm run dev` runs the dashboard against the same Apps Script backend as production
(there's no separate local backend) — so steps 1–2 above need to be done before `npm run
dev` will actually load data, though the login screen itself will render either way.
