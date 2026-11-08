# Cabin TV Tracker – mobile app

## Tracker web app source (`apps-script/`)
The Combined PAX + Crew TV Installation Tracker (Google Apps Script):
- `apps-script/Code.gs` – server logic + cabin inventory
- `apps-script/Index.html` – the web app page
- `apps-script/PasswordSetup.html` – technician password setup dialog
- `Photo.html` (header photo) is optional and not stored here.

The code has **no built-in cabins, crew offices or public areas**. The
tracker starts empty and the admin uploads their own lists.

### Set up your own tracker
1. Create a Google Sheet → Extensions → Apps Script. Add `Code.gs`,
   `Index.html` and `PasswordSetup.html` (same file names).
2. Run `setup()` once from the editor and allow the permissions.
3. Project Settings → Script Properties: add `TRACKER_ADMIN_SETUP_KEY`
   (any long secret). Reload the sheet and use the **Tracker Admin →
   Set / Reset Technician Passwords** menu to set the Admin password.
   Accounts: `Admin` (signs in with this password) and `TV Check 1`–`TV Check 6`
   (sign in by choosing their name, no password). Change `TECHNICIANS` /
   `PASSWORD_ACCOUNTS` in Code.gs (and `TECHS` / `PASSWORD_ACCOUNTS` in Index.html)
   to add accounts or require more passwords.
4. Deploy → New deployment → Web app (Execute as: Me). Open the `/exec` link
   and sign in as the admin.
5. Open **Upload Cabins & Public Areas**, download the templates, fill them
   in and upload them (.xlsx, .xls or .csv):
   - Cabins / crew offices: `Cabin Type` (PAX or CREW), `Cabin No`, `Deck`, `FZ`, `TV Count`
   - Public areas: `Venue`, `Location`, `Deck`, `TV Size`, `QTY`, `Area` (PAX or CREW), `Source`

   "Add new and update existing" keeps all ticks, remarks and assignments.
   "Replace the whole list" deletes rows that are not in the file.
   "🗑 Remove all cabins" / "🗑 Remove all public areas" empties a list (type
   DELETE to confirm); a backup copy of the spreadsheet is saved to Drive first.

Passwords, the admin setup key and the phone-app API key live in the
project's Script Properties, never in these files.

## Run it on Cloudflare instead of Google (`cloudflare/`)
The same tracker can run on Cloudflare Workers with no Google account and no
Google Sheet. `cloudflare/build.mjs` packs `apps-script/Code.gs` and `Index.html`
unchanged into one Worker, with stand-ins for the Google services it uses. All
data lives in a Durable Object with built-in SQLite storage (works on the
Workers **Free** plan). Change the tracker in `apps-script/` and both versions
get it.

**One-time setup (about 10 minutes, all in the browser):**
1. Cloudflare dashboard → **My Profile → API Tokens → Create Token** → use the
   **"Edit Cloudflare Workers"** template → Create. Copy the token.
2. Cloudflare dashboard → **Workers & Pages**: copy your **Account ID** (right side).
3. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**, add:
   - `CLOUDFLARE_API_TOKEN` – the token from step 1
   - `CLOUDFLARE_ACCOUNT_ID` – the ID from step 2
   - `ADMIN_SETUP_KEY` – any long secret (you type it when setting the Admin password)
4. GitHub repo → **Actions → Deploy tracker to Cloudflare → Run workflow**.
   It also runs by itself whenever `apps-script/` or `cloudflare/` changes.
5. Open `https://cabin-tv-tracker.<your-subdomain>.workers.dev/setup`, enter the
   setup key and set the **Admin** password.
6. Open `https://cabin-tv-tracker.<your-subdomain>.workers.dev/`, sign in as Admin
   and upload your cabins and public areas.

**What is different from the Google version**
- No spreadsheet: view and change data in the web app, export with **Export All
  to Excel**, and load lists with **Upload Cabins & Public Areas**.
- The Admin password is set at `/setup` (instead of the spreadsheet menu).
- Backups are stored inside the Worker and download as .xlsx from the link the
  backup shows.
- Emails (daily summary, sign-in notices) need a free [Resend](https://resend.com)
  account: add the Worker secrets `RESEND_API_KEY` and `MAIL_FROM`. Without them
  the tracker works and simply sends no email.
- Time zone: `TIMEZONE` in `cloudflare/wrangler.toml` (default `UTC`).
- Phone app: set `API_URL` in `App.js` to the workers.dev link (ending in `/`).

Test locally (Node 22): `cd cloudflare && npm test`.

### Open it on iPhone and Android
The web app works in the phone's browser; no app store install is needed.
Share your `/exec` link with the team, then:
- **iPhone:** open the link in **Safari** → Share button → **Add to Home Screen** → Add.
- **Android:** open the link in **Chrome** → ⋮ menu → **Add to Home screen** (or **Install app**) → Add.

Tips:
- Set the deployment's **Who has access** to **Anyone**, so technicians sign in with their
  tracker username and password and don't need a Google account.
- After changing the code, use **Deploy → Manage deployments → Edit → New version**. The
  `/exec` link stays the same, so home-screen icons keep working.
- If an iPhone shows "page not found" (several Google accounts signed in), open the link
  in a Private tab in Safari.
- Phones need internet: the web app runs on Google's servers.

Native Android + iOS app (Expo / React Native) that reads your
PAX + Crew TV Installation Tracker through your Apps Script.

## 1. Apps Script (once)
1. Add `apps-script/Api.gs` to your Apps Script project.
2. Set `API_KEY` in Api.gs to a long secret. Set `API_SHEET_ID` only if the script isn't bound to the sheet.
3. At the top of your existing `doGet(e)`, add:
   `if (e && e.parameter && e.parameter.api) return handleApi_(e);`
4. Deploy → Manage deployments → Edit → New version. Access: "Anyone".
5. Test in a browser: `<your /exec URL>?api=sheets&key=<your key>` — you should see JSON.

## 2. App settings
In `App.js`, set `API_KEY` to the same secret (URL is already filled in).

## 3. Try it on your phone
    npm install
    npx expo install expo react react-native
    npx expo start
Scan the QR code with the **Expo Go** app (Android / iPhone).

## 4. Build installable apps
    npm install -g eas-cli
    eas login
    eas build -p android --profile preview      # APK to install directly
    eas build -p android --profile production   # .aab for Google Play
    eas build -p ios --profile production       # needs Apple Developer account
    eas submit -p android   /   eas submit -p ios

## Build the APK on GitHub (no computer needed)
1. Create a new repo on github.com (can be private).
2. Upload everything in this folder, including the `.github` folder.
3. Open the **Actions** tab → "Build Android APK" runs automatically (about 10–15 min).
4. When it shows a green tick, open the run → **Artifacts** → download `CabinTVTracker-apk`.
5. Unzip it on your Android phone and tap the `.apk` to install (allow "install unknown apps").
Every time you change a file, a new APK is built.
