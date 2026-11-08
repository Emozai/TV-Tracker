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
   Set / Reset Technician Passwords** menu to give each account a password.
   Account names are in `TECHNICIANS` at the top of `Code.gs`; the admin
   account is `Admin`.
   Only `Admin` needs a password; technicians sign in by choosing their name
   (change `PASSWORD_ACCOUNTS` in Code.gs and Index.html to require more).
4. Deploy → New deployment → Web app (Execute as: Me). Open the `/exec` link
   and sign in as the admin.
5. Open **Upload Cabins & Public Areas**, download the templates, fill them
   in and upload them (.xlsx, .xls or .csv):
   - Cabins / crew offices: `Cabin Type` (PAX or CREW), `Cabin No`, `Deck`, `FZ`, `TV Count`
   - Public areas: `Venue`, `Location`, `Deck`, `TV Size`, `QTY`, `Area` (PAX or CREW), `Source`

   "Add new and update existing" keeps all ticks, remarks and assignments.
   "Replace the whole list" deletes rows that are not in the file.

Passwords, the admin setup key and the phone-app API key live in the
project's Script Properties, never in these files.

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
