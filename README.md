# Sahoo Finance Ledger

A loan & collection management app, shared across your owner login and all your
agents. Everyone sees and updates the same data, no matter which laptop or phone
they're using.

Data is stored in a **Google Sheet in your Drive** (via a small Google Apps
Script "bridge"), so it's shared, and lives in an account you already control.

## Step 1 - Set up the shared Google Sheet (do this once)

1. Go to https://sheets.new to create a fresh Google Sheet. Name it anything,
   e.g. "Sahoo Finance Data".
2. In the Sheet, click **Extensions → Apps Script**.
3. Delete any code shown in the editor, then paste in the entire contents of
   **`google-apps-script/Code.gs`** (included in this project).
4. Click **Deploy → New deployment**. For "Select type", choose **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**. Google will ask you to authorize it — approve it (it's
   your own script acting on your own Sheet).
6. Copy the **Web app URL** shown (it ends in `/exec`). This is your shared
   data link.

## Step 2 - Connect the app to your Sheet

1. Open **`src/App.jsx`** in this project (any text editor works, e.g. Notepad,
   VS Code, or even editing the file directly on GitHub).
2. Find this line near the top:
   ```js
   const SYNC_URL = "PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE";
   ```
3. Replace the placeholder text with the URL you copied in Step 1, keeping the
   quote marks, e.g.:
   ```js
   const SYNC_URL = "https://script.google.com/macros/s/AKfycb.../exec";
   ```
4. Save the file.

## Step 3 - Put the app online so everyone can reach it

The easiest free option is **Vercel**:

1. Create a free account at https://vercel.com and a free account at
   https://github.com if you don't already have them.
2. Create a new empty repository on GitHub, then upload every file in this
   project folder to it (GitHub's website lets you drag-and-drop files
   straight into the browser — no command line required). Make sure you
   upload the version of `src/App.jsx` with your Sheet's URL already pasted in.
3. In Vercel, click **"Add New Project" → "Import"** and choose that GitHub
   repository.
4. Vercel detects it's a Vite app automatically — click **Deploy**.
5. After about a minute, Vercel gives you a permanent link, e.g.
   `https://sahoo-finance-ledger.vercel.app`.

Share that one link with every agent and with the owner. Whoever opens it, on
whichever laptop or phone, sees the same live data — and any payment, new
client, or new loan one person enters shows up for everyone else within about
15 seconds (or immediately if they tap the refresh icon next to "Synced" in
the sidebar).

## (Optional) Preview it on your own computer first

1. Install [Node.js](https://nodejs.org) (the LTS version) if you don't have it.
2. Open a terminal inside this project folder and run:
   ```
   npm install
   npm run dev
   ```
3. Open the link it prints (usually `http://localhost:5173`).

To check it on your phone during this local preview: make sure your phone is
on the same WiFi as your laptop, run `npm run dev -- --host` instead, and
open the "Network" address it prints on your phone's browser.

## Notes

- The sidebar shows a small dot and "Synced …" text: green means it's reading
  and writing the shared Google Sheet; amber means `SYNC_URL` hasn't been set
  yet, so it's only saving to this one device.
- The **"Backup by email"** / **"Restore backup"** buttons still work as
  before — handy for an extra copy, even though the Sheet is now the main
  shared source of truth.
- If you ever want to reset to sample demo data, use "Reset demo data" in the
  sidebar (owner login only) — note this clears the shared data for everyone.
- You can open the Google Sheet directly any time to see when it was last
  written to, though the data itself is stored as one long technical block of
  text in cell A1 rather than in readable columns — that's expected.
