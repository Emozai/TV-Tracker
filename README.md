# Escape TV Tracker – mobile app

## Tracker web app source (`apps-script/`)
The NCL Escape Combined PAX + Crew TV Installation Tracker (Google Apps Script):
- `apps-script/Code.gs` – server logic + cabin inventory
- `apps-script/Index.html` – the web app page
- `apps-script/PasswordSetup.html` – technician password setup dialog
- `Photo.html` (header photo) is optional and not stored here.

Passwords, the admin setup key and the phone-app API key live in the
project's Script Properties, never in these files. To restore, paste each
file into an Apps Script project with the same name, then Deploy → Web app.

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
4. When it shows a green tick, open the run → **Artifacts** → download `EscapeTVTracker-apk`.
5. Unzip it on your Android phone and tap the `.apk` to install (allow "install unknown apps").
Every time you change a file, a new APK is built.
