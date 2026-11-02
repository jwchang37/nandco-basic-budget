# Estimator backend setup (about 5 minutes)

The website is a static page, so it uses a small Google Apps Script attached to a Google Sheet to
store estimate requests, email them to hello@nandcoconstruction.com, and hold the global
cost multiplier / GC fee. It's free and runs under your Google account.

## 1. Create the sheet and script
1. Go to https://sheets.new and name the sheet something like **NANDCO Estimate Requests**.
2. In the sheet: **Extensions → Apps Script**.
3. Delete the starter code, paste in everything from [`Code.gs`](Code.gs), and click **Save**.

## 2. Set the admin password
1. In the Apps Script editor: **Project Settings** (gear icon) → **Script Properties** → **Add script property**.
2. Property: `ADMIN_KEY`, Value: the password you'll type after clicking **NANDCO** on the site. Save.
   (Pick something new and long. It never goes in the website code.)

## 3. Authorize
1. Back in the **Editor**, choose the `authorize` function in the toolbar dropdown and click **Run**.
2. Google asks for permission (sheet access + send email as you). Click through
   **Advanced → Go to project (unsafe) → Allow**. It says "unsafe" only because the script is yours and unverified.

## 4. Deploy as a web app
1. **Deploy → New deployment** → gear icon → **Web app**.
2. Execute as: **Me**. Who has access: **Anyone**.
3. Click **Deploy** and copy the **Web app URL** (ends in `/exec`).

## 5. Connect the website
Paste the URL into `index.html`:

```js
const BACKEND_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

(or send the URL to Claude and it'll make the change).

## Updating the script later
The website and the script are separate. Changing `index.html` never requires touching the script.
Only if `Code.gs` itself changes: **Deploy → Manage deployments → ✏️ (edit) → Version: New version → Deploy**.
That keeps the same URL. Don't use **New deployment** again, because that makes a new URL.

## Notes
- Emails come from the Google account that deployed the script, with Reply-To set to the customer.
  Consumer Gmail accounts can send about 100 of these a day; Google Workspace about 1,500.
- Every request is a row in the **Requests** tab; you can view or edit it directly.
- Global multiplier / GC fee are stored as Script Properties `MULTIPLIER_PCT` and `GC_FEE_PCT`
  (default 20 / 15 until you save from the admin portal).
