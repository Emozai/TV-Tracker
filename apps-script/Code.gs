/** COMBINED PAX + CREW TV INSTALLATION TRACKER (Google Apps Script web app)
 *
 * Apps Script project files:
 *   Code.gs            – server logic
 *   Index.html         – the web app page
 *   PasswordSetup.html – technician password dialog (spreadsheet menu)
 *   Photo.html         – header photo as a data URL (optional)
 *
 * NO BUILT-IN SHIP DATA (2026-11-08)
 *   This version has no cabin, crew office or public area list in the code. The tracker starts empty and
 *   the admin uploads their own lists in the web app: sign in as the admin › "Upload Cabins & Public Areas".
 *     - Cabins and crew offices: .xlsx, .xls or .csv with the columns
 *         Cabin Type (PAX or CREW) · Cabin No · Deck · FZ · TV Count     (only Cabin No is required)
 *       Crew cabins and crew offices use Cabin Type CREW. A crew number that starts with a digit gets a "C"
 *       in front (1234 -> C1234), so passenger and crew numbers never clash.
 *     - Public areas: Venue · Location · Deck · TV Size · QTY · Area (PAX or CREW) · Source
 *       (Venue or Location is required).
 *     - "Add new and update existing" adds new rows and updates Deck / FZ / TV Count (or Location details)
 *       of rows already there; checklist ticks, remarks and assignments are kept.
 *       "Replace the whole list" makes the sheet match the file exactly: rows not in the file are removed
 *       (with their ticks and remarks); rows in both keep everything.
 *     - Each upload page has a "Download template" button with the right column names.
 *   Server functions: importCabinInventory / importVenueInventory (admin only).
 *
 * FIRST-TIME SETUP
 *   1. Create a Google Sheet, open Extensions › Apps Script, and add the files above.
 *   2. Run setup() once (creates the tracker sheet, turns on live sheet sync and the email / backup jobs).
 *   3. Script Properties: add TRACKER_ADMIN_SETUP_KEY (any long secret), then reload the sheet and use
 *      Tracker Admin › "Set / Reset Technician Passwords" to give each account a password.
 *   4. Deploy › New deployment › Web app (Execute as: Me). Open the /exec link and sign in as the admin.
 *   5. Upload your cabins and public areas.
 *   Technician account names are set in TECHNICIANS below; the admin account is "Admin".
 *
 * LIVE SHEET SYNC
 *   - The "Combined Cabin TV Checklist" sheet is the source of truth. Add, delete,
 *     rename or re-order cabins, or change Cabin Type, Deck, FZ, TV Count, checklist
 *     values, Assigned To or Comments directly in the sheet: the web app picks the change
 *     up within a few seconds. No redeploy is needed for sheet changes.
 *   - setup() never wipes the sheet. Only an upload with "Replace the whole list" removes rows.
 *   - Saves are matched by cabin type + cabin number, not by row number.
 */

// ---------------------------------------------------------------------------
// CONFIGURATION
// ---------------------------------------------------------------------------
const SHEET_NAME = 'Combined Cabin TV Checklist';
const TV_STATUS_LABELS = ['TV installed','Power Connected','No TV Here','Tested','Fully Configured','Held Up'];
const CHECK_FIELDS = TV_STATUS_LABELS.slice();
const COMMON_CHECK_FIELDS = [];
const SECURITY_DOOR = 'Security to Open the Door'; // PAX only, one per cabin (shown on the TV 1 checklist). Does not affect status.
const CABIN_EXTRA_FIELDS = ['Patch cable','AP', SECURITY_DOOR]; // Patch cable: PAX "Patch Cable" / CREW "LAN Connected". AP + Security: PAX only.

function activeCheckFieldsForType_(type) {
  return String(type || '').trim().toUpperCase() === 'PAX'
    ? ['No TV Here','TV installed','Power Connected','Patch cable','AP','Fully Configured','Tested','Held Up']
    : ['No TV Here','TV installed','Power Connected','Patch cable','Fully Configured','Tested','Held Up']; // CREW: Patch cable column = "LAN Connected"
}

function tvFieldName_(tvNumber, label) {
  if (Number(tvNumber) === 1) return label;
  const suffix = ({'TV installed':'Installed'})[label] || label;
  return 'TV ' + tvNumber + ' ' + suffix;
}

const TV2_FIELDS = TV_STATUS_LABELS.map(label => tvFieldName_(2, label));
const TV3_FIELDS = TV_STATUS_LABELS.map(label => tvFieldName_(3, label));
// TV 2 / TV 3 get a Patch cable (PAX) / LAN Connected (CREW) check but NO AP (AP is PAX TV 1 only).
const PAX_EXTRA_TV_FIELDS = [tvFieldName_(2, 'Patch cable'), tvFieldName_(3, 'Patch cable')];
const TV1_FIELDS = TV_STATUS_LABELS.slice();
const ALL_TV_FIELDS = TV1_FIELDS.concat(TV2_FIELDS, TV3_FIELDS, PAX_EXTRA_TV_FIELDS);

const HEADERS = ['CABIN TYPE','CABIN NO','DECK','FZ','TV installed','Power Connected','No TV Here','Patch cable','AP',SECURITY_DOOR,
  'Tested','Fully Configured','Held Up','Check By','Assigned To','Status','Checked At','TV Count','Comments','What to Check','Locked']
  .concat(TV2_FIELDS, TV3_FIELDS, PAX_EXTRA_TV_FIELDS);

const TV_CHECK_ACCOUNTS = ['TV Check 1','TV Check 2','TV Check 3','TV Check 4','TV Check 5','TV Check 6'];
const TECHNICIANS = ['Tech 1','Tech 2','Tech 3','Tech 4','Admin'].concat(TV_CHECK_ACCOUNTS);
const AUTHORIZED_TECHNICIANS = TECHNICIANS.slice();
// Only these accounts need a password to sign in. Every other account signs in by choosing its name.
const PASSWORD_ACCOUNTS = ['Admin'];
function needsPassword_(name) { return PASSWORD_ACCOUNTS.includes(accountName_(name)); }
const AUTH_CACHE_TTL_SECONDS = 21600; // 6 hours
const EMAIL_QUEUE_SHEET = 'Tracker Email Queue';
const ACTIVE_USERS_SHEET = 'Tracker Active Users';
const AUDIT_SHEET = 'Tracker Update History';
const TRACKER_EMAIL_PROPERTY = 'TRACKER_NOTIFICATION_EMAIL';
const DATA_VERSION_PROPERTY = 'TRACKER_DATA_VERSION';

/** Reads a TV Count cell: 2, "2", "2 TVs", "TV x2", "two" -> 2. Returns 0 when blank or unreadable. */
function parseTvCount_(v) {
  if (typeof v === 'number') return v >= 1 && v <= 3 ? Math.round(v) : 0;
  const s = normalize_(v);
  if (!s) return 0;
  const words = { one: 1, single: 1, two: 2, double: 2, three: 3, triple: 3 };
  const w = Object.keys(words).find(k => new RegExp('\\b' + k + '\\b').test(s));
  if (w) return words[w];
  const d = s.match(/[1-3]/);
  return d ? Number(d[0]) : 0;
}

/** Fallback TV count, only used when the sheet's TV Count cell is blank: 1 TV. */
function getTvCount_(type, cabin) { return 1; }

// ---------------------------------------------------------------------------
// WEB APP ENTRY
// ---------------------------------------------------------------------------
function doGet(e) {
  const page = e && e.parameter ? String(e.parameter.page || '').toLowerCase() : '';
  if (page === 'keyrunner' && typeof keyRunnerPage_ === 'function') return keyRunnerPage_();     // optional KeyRunner.gs
  if (page === 'finalcheck' && typeof finalCheckPage_ === 'function') return finalCheckPage_(); // optional FinalCheck.gs
  if (e && e.parameter && e.parameter.api) return handleApi_(e);   // mobile app JSON API
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Combined Cabin TV Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Returns the content of another HTML file (used for the header photo). Missing file -> ''. */
function include(name) {
  try { return HtmlService.createHtmlOutputFromFile(name).getContent().trim(); }
  catch (e) { return ''; }
}

function onOpen() {
  try { if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VENUE_SHEET)) ensureVenueSheet_(); } catch (err) { console.warn('Venue check on open skipped: ' + err); }
  const menu = SpreadsheetApp.getUi().createMenu('Tracker Admin')
    .addItem('Set / Reset Technician Passwords', 'openTechnicianPasswordSetup')
    .addItem('Turn on live sheet sync (run once)', 'enableLiveSyncFromMenu')
    .addItem('Recalculate all statuses', 'recalcAllStatusesFromMenu')
    .addItem('Crew: sync TV Installed / Power / LAN Connected', 'syncCrewLanConnectedFromMenu')
    .addItem('Set up / check Public Area TVs sheet', 'setupVenueSheetFromMenu');
  if (typeof setupFinalCheckFromMenu === 'function') menu.addItem('Set up / check Final Check sheet', 'setupFinalCheckFromMenu');
  menu.addSeparator()
    .addItem('Turn on nightly backup, clean-up and daily progress', 'installMaintenanceFromMenu')
    .addItem('Back up spreadsheet now', 'backupNowFromMenu')
    .addItem('Set daily summary email…', 'setSummaryEmailFromMenu')
    .addItem('Send daily summary now', 'sendSummaryNowFromMenu')
    .addItem('Show phone app API key', 'showApiKeyFromMenu');
  if (typeof setKeyRunnerPasswordFromMenu === 'function') menu.addItem('Set key runner password', 'setKeyRunnerPasswordFromMenu');
  menu.addItem('Create / reset TV Check passwords (TV Check 1–6)', 'createTvCheckPasswordsFromMenu')
    .addToUi();
}

/** Built-in copy of the password dialog, used when the project has no PasswordSetup.html file. */
const PASSWORD_SETUP_HTML_ = "<!doctype html><html><head><base target=\"_top\"><meta charset=\"utf-8\"><style>body{font:14px Arial,sans-serif;padding:18px;background:#0b1025;color:#f8fafc}h2{margin:0 0 8px;color:#67e8f9}p{color:#cbd5e1;line-height:1.4}.field{display:block;margin:12px 0 5px;font-weight:bold;color:#c4b5fd}input,select{box-sizing:border-box;width:100%;padding:10px;border-radius:8px;border:1px solid #475569;background:#111827;color:white}button{margin-top:14px;padding:10px 14px;border:0;border-radius:8px;color:white;background:linear-gradient(100deg,#7c3aed,#db2777,#0891b2);font-weight:bold;cursor:pointer}#result{margin-top:12px;white-space:pre-wrap;line-height:1.4}.note{font-size:12px}</style></head><body><h2>Technician Password Setup</h2><p>Set or reset the password for one account. Passwords are stored as salted hashes, not plaintext.</p><label class=\"field\" for=\"adminKey\">Admin setup key</label><input id=\"adminKey\" type=\"password\" autocomplete=\"off\" placeholder=\"Enter TRACKER_ADMIN_SETUP_KEY\"><label class=\"field\" for=\"username\">Technician account</label><select id=\"username\">__OPTIONS__</select><label class=\"field\" for=\"password\">New password (minimum 12 characters)</label><input id=\"password\" type=\"password\" autocomplete=\"new-password\"><label class=\"field\" for=\"confirm\">Confirm password</label><input id=\"confirm\" type=\"password\" autocomplete=\"new-password\"><button id=\"save\" onclick=\"savePassword()\">Save Password</button><div id=\"result\" role=\"status\"></div><p class=\"note\">Keep the admin setup key private. Do not share it with technicians.</p><script>function savePassword(){const key=document.getElementById('adminKey').value,name=document.getElementById('username').value,p=document.getElementById('password').value,c=document.getElementById('confirm').value,out=document.getElementById('result'),b=document.getElementById('save');if(p.length<12){out.textContent='Password must be at least 12 characters.';return;}if(p!==c){out.textContent='Passwords do not match.';return;}b.disabled=true;out.textContent='Saving password…';google.script.run.withSuccessHandler(r=>{b.disabled=false;out.textContent=(r&&r.message)||'Saved.';document.getElementById('password').value='';document.getElementById('confirm').value='';}).withFailureHandler(e=>{b.disabled=false;out.textContent='Error: '+(e.message||e);}).saveTechnicianPasswordSetup(key,name,p);}</script></body></html>";

function openTechnicianPasswordSetup() {
  let html;
  try { html = HtmlService.createHtmlOutputFromFile('PasswordSetup'); }
  catch (e) {
    const opts = TECHNICIANS.map(n => '<option>' + htmlEsc_(n) + '</option>').join('');
    html = HtmlService.createHtmlOutput(PASSWORD_SETUP_HTML_.replace('__OPTIONS__', opts));
  }
  SpreadsheetApp.getUi().showModalDialog(html.setWidth(440).setHeight(520), 'TV Tracker — Technician Password Setup');
}

// ---------------------------------------------------------------------------
// HELPERS
// ---------------------------------------------------------------------------
function normalize_(v) { return String(v == null ? '' : v).replace(/\u00a0/g, ' ').trim().replace(/\s+/g, ' ').toLowerCase(); }
function normalizeYesNo_(v) { return v === true || ['yes','true','1','y','checked'].includes(normalize_(v)) ? 'Yes' : 'No'; }
function formatDate_(v) {
  if (!v) return '';
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : String(v);
}

/** PAX / CREW from the sheet value; blank or unusual values fall back to the cabin number (C… = CREW). */
function inferType_(type, cabin) {
  const t = String(type == null ? '' : type).trim().toUpperCase();
  if (t === 'PAX' || t === 'CREW') return t;
  if (t === 'PASSENGER' || t === 'GUEST') return 'PAX';
  return /^C\d/i.test(String(cabin == null ? '' : cabin).trim()) ? 'CREW' : 'PAX';
}
function cabinKey_(type, cabin) { return inferType_(type, cabin) + ':' + String(cabin == null ? '' : cabin).trim().toUpperCase(); }
/** "tech 1 " typed in the sheet -> "Tech 1". Unknown names are kept as typed. */
function techName_(v) { const s = String(v == null ? '' : v).trim(); return accountName_(s) || s; }

/** Bumped on every change; the web app polls this cheap value and reloads only when it moves. */
function bumpDataVersion_() {
  try { PropertiesService.getScriptProperties().setProperty(DATA_VERSION_PROPERTY, String(Date.now())); }
  catch (e) { console.warn('Data version not updated: ' + e); }
}
function getDataVersion() {
  return PropertiesService.getScriptProperties().getProperty(DATA_VERSION_PROPERTY) || '0';
}

/** Status is based on the TV 1 checklist: Fully Configured = Yes -> Fully Configured; any Yes -> In Progress.
 *  "Security to Open the Door" is a flag only and is not part of the status. */
function calculateCabinStatusFromFields_(readField, tvCount, type) {
  let allConfigured = true, any = false;
  activeCheckFieldsForType_(type).forEach(field => {
    const yes = normalizeYesNo_(readField(field)) === 'Yes';
    if (yes) any = true;
    if (field === 'Fully Configured' && !yes) allConfigured = false;
  });
  return allConfigured ? 'Fully Configured' : any ? 'In Progress' : 'Not Started';
}

function canManageAssignments_(username) { return ['Tech 3','Admin'].includes(accountName_(username)); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script project from the Google Spreadsheet you want to use.');
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

/** Adds any missing columns in place and fills blank TV Count / No TV Here cells. Never overwrites typed values. */
function ensureTrackerColumns_(sh) {
  if (!sh || sh.getLastRow() < 1) return;
  let lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const known = {};
  headers.forEach((h, i) => { if (normalize_(h)) known[normalize_(h)] = i; });

  let added = false;
  ['TV Count','Comments','Tested','Held Up','No TV Here','Patch cable','AP','What to Check','Locked',SECURITY_DOOR]
    .concat(TV2_FIELDS, TV3_FIELDS, PAX_EXTRA_TV_FIELDS)
    .forEach(name => {
      if (known[normalize_(name)] === undefined) {
        added = true;
        lastCol++;
        sh.getRange(1, lastCol).setValue(name).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
        known[normalize_(name)] = lastCol - 1;
      }
    });

  // The blank TV Count / No TV Here fill below reads the whole sheet, so it runs at most once every
  // 5 minutes (or straight away when a column was just added). Blank cells are handled in getData_ anyway.
  const cache = CacheService.getScriptCache(), fillKey = 'TRACKER_BLANK_FILL_DONE';
  if (!added && cache.get(fillKey)) return;
  try { cache.put(fillKey, '1', 300); } catch (_) {}

  const cabinCol = known[normalize_('CABIN NO')];
  const typeCol = known[normalize_('CABIN TYPE')];
  const countCol = known[normalize_('TV Count')];
  const noTvCol = known[normalize_('No TV Here')];
  const lastRow = sh.getLastRow();
  if (lastRow < 2 || cabinCol === undefined) return;

  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const countValues = [], noTvValues = [];
  let countsNeedWrite = false, noTvNeedsWrite = false;
  data.forEach(row => {
    const cabin = String(row[cabinCol] == null ? '' : row[cabinCol]).trim();
    const type = inferType_(typeCol === undefined ? '' : row[typeCol], cabin);
    let count = row[countCol];
    // Only fill a truly blank cell; anything typed (even "2 TVs") is left exactly as the user wrote it.
    if (cabin && (count === '' || count == null)) { count = getTvCount_(type, cabin); countsNeedWrite = true; }
    countValues.push([count == null ? '' : count]);
    let noTv = row[noTvCol];
    if (cabin && (noTv === '' || noTv == null)) { noTv = 'No'; noTvNeedsWrite = true; }
    noTvValues.push([noTv == null ? '' : noTv]);
  });
  if (countsNeedWrite) sh.getRange(2, countCol + 1, countValues.length, 1).setValues(countValues);
  if (noTvNeedsWrite) sh.getRange(2, noTvCol + 1, noTvValues.length, 1).setValues(noTvValues);
}

function headerMap_(sh) {
  ensureTrackerColumns_(sh);
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0], m = {};
  h.forEach((v, i) => m[normalize_(v)] = i);
  HEADERS.forEach(k => { if (m[normalize_(k)] === undefined) throw new Error('Missing column ' + k + '. Run setup() first.'); });
  return m;
}

/** Finds a cabin's current row. Tries the row the browser knew first, then searches the sheet. 0 = not found. */
function resolveCabinRow_(sh, m, type, cabin, hintRow) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')];
  const hasCabin = String(cabin == null ? '' : cabin).trim() !== '';
  const hintOk = Number.isInteger(hintRow) && hintRow >= 2 && hintRow <= last;
  if (!hasCabin) return hintOk ? hintRow : 0; // older pending saves without a cabin number
  const want = cabinKey_(type, cabin), width = Math.max(ci, ti === undefined ? 0 : ti) + 1;
  const keyAt = r => String(r[ci] == null ? '' : r[ci]).trim() ? cabinKey_(ti === undefined ? '' : r[ti], r[ci]) : '';
  if (hintOk && keyAt(sh.getRange(hintRow, 1, 1, width).getValues()[0]) === want) return hintRow;
  const vals = sh.getRange(2, 1, last - 1, width).getValues();
  for (let i = 0; i < vals.length; i++) if (keyAt(vals[i]) === want) return i + 2;
  return 0;
}

/** Rewrites the Status column for the given rows from their checklist values. */
function recalcStatusRows_(sh, startRow, numRows) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  numRows = Math.min(numRows, lastRow - startRow + 1);
  if (startRow < 2 || numRows < 1 || lastCol < 1) return;
  const m = {};
  sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach((h, i) => { if (normalize_(h)) m[normalize_(h)] = i; });
  const col = n => m[normalize_(n)];
  const sCol = col('Status'), cCol = col('CABIN NO'), tCol = col('CABIN TYPE');
  if (sCol === undefined || cCol === undefined) return;
  const rows = sh.getRange(startRow, 1, numRows, lastCol).getValues();
  const out = rows.map(r => {
    const cabin = String(r[cCol] == null ? '' : r[cCol]).trim();
    if (!cabin) return [r[sCol]];
    const type = inferType_(tCol === undefined ? '' : r[tCol], cabin);
    return [calculateCabinStatusFromFields_(f => { const i = col(f); return i === undefined ? 'No' : r[i]; }, 1, type)];
  });
  sh.getRange(startRow, sCol + 1, numRows, 1).setValues(out);
}

// ---------------------------------------------------------------------------
// CREW: TV INSTALLED = POWER CONNECTED = LAN CONNECTED (2026-10-19)
// On crew cabins these three checks are always the same, for TV 1, TV 2 and TV 3.
// ---------------------------------------------------------------------------
const CREW_LINKED_LABELS = ['TV installed','Power Connected','Patch cable']; // Patch cable = "LAN Connected" on crew
const CREW_LAN_SYNC_PROPERTY = 'TRACKER_CREW_LAN_SYNC_20261019';

function crewLinkedFields_(n) { return CREW_LINKED_LABELS.map(l => tvFieldName_(n, l)); }

/** One-time clean-up: on every crew cabin, if TV installed, Power Connected or LAN Connected is Yes, all three become Yes. */
function syncCrewLanConnected() {
  const sh = sheet_(), m = headerMap_(sh), last = sh.getLastRow();
  if (last < 2) return 'The tracker sheet is empty.';
  const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')];
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const touched = {}, fixed = [];
  rows.forEach(r => {
    const cabin = String(r[ci] == null ? '' : r[ci]).trim();
    if (!cabin || inferType_(r[ti], cabin) !== 'CREW') return;
    let changed = false;
    for (let n = 1; n <= 3; n++) {
      const cols = crewLinkedFields_(n).map(f => m[normalize_(f)]);
      if (cols.some(c => c === undefined)) continue;
      const yesCol = cols.find(c => normalizeYesNo_(r[c]) === 'Yes');
      if (yesCol === undefined) continue;
      cols.forEach(c => {
        if (normalizeYesNo_(r[c]) !== 'Yes') { r[c] = r[yesCol] === true ? true : 'Yes'; touched[c] = true; changed = true; }
      });
    }
    if (changed) fixed.push(cabin);
  });
  Object.keys(touched).map(Number).forEach(c => sh.getRange(2, c + 1, rows.length, 1).setValues(rows.map(r => [r[c]])));
  if (fixed.length) {
    recalcStatusRows_(sh, 2, rows.length);
    try {
      const now = new Date(), h = ensureActivitySheets_().history;
      const audit = fixed.map(c => [now, 'Script', 'CREW', c, 'Synced TV Installed / Power Connected / LAN Connected', '', '']);
      h.getRange(h.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
    } catch (_) {}
    bumpDataVersion_();
  }
  PropertiesService.getScriptProperties().setProperty(CREW_LAN_SYNC_PROPERTY, new Date().toISOString());
  SpreadsheetApp.flush();
  const msg = 'Crew cabins synced (TV Installed = Power Connected = LAN Connected): ' + fixed.length + ' cabin(s) updated.';
  Logger.log(msg);
  return msg;
}

/** Runs the clean-up once, automatically, the first time data is read after this update. */
function ensureCrewLanSynced_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(CREW_LAN_SYNC_PROPERTY)) return;
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return;
    try { if (!props.getProperty(CREW_LAN_SYNC_PROPERTY)) syncCrewLanConnected(); }
    finally { lock.releaseLock(); }
  } catch (err) { Logger.log('Crew LAN sync skipped: ' + err); }
}

function syncCrewLanConnectedFromMenu() { SpreadsheetApp.getUi().alert(syncCrewLanConnected()); }

/** Hand edits in the sheet: changing one of the three cells on a crew row copies that value to the other two. */
function syncCrewLinkedOnEdit_(sh, range) {
  const lastCol = sh.getLastColumn(), m = {};
  sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach((h, i) => { if (normalize_(h)) m[normalize_(h)] = i; });
  const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')];
  if (ci === undefined) return;
  const c1 = range.getColumn() - 1, c2 = range.getLastColumn() - 1;
  const inRange = c => c >= c1 && c <= c2;
  const groups = [1, 2, 3].map(n => crewLinkedFields_(n).map(f => m[normalize_(f)]))
    .filter(g => g.every(c => c !== undefined) && g.some(inRange));
  if (!groups.length) return;
  const start = Math.max(2, range.getRow()), num = range.getLastRow() - start + 1;
  if (num < 1) return;
  const rows = sh.getRange(start, 1, num, lastCol).getValues(), touched = {};
  rows.forEach(r => {
    const cabin = String(r[ci] == null ? '' : r[ci]).trim();
    if (!cabin || inferType_(ti === undefined ? '' : r[ti], cabin) !== 'CREW') return;
    groups.forEach(g => {
      const src = g.find(inRange);
      g.forEach(c => { if (c !== src && normalizeYesNo_(r[c]) !== normalizeYesNo_(r[src])) { r[c] = r[src]; touched[c] = true; } });
    });
  });
  Object.keys(touched).map(Number).forEach(c => sh.getRange(start, c + 1, num, 1).setValues(rows.map(r => [r[c]])));
}

// ---------------------------------------------------------------------------
// PUBLIC AREA TVs — Crew & Pax areas (venues), 2026-10-19
// Separate sheet "Public Area TVs". One row = one TV location (venue, location, deck, TV size, qty, area, source),
// uploaded by the admin in the web app or typed into the sheet. The checks on a row apply to all QTY TVs on it.
// The sheet is the source of truth: add, delete or edit rows there and the web app follows.
// ---------------------------------------------------------------------------
const VENUE_SHEET = 'Public Area TVs';
const VENUE_CHECKS = ['Installed','Power Connected','LAN Connected','Powering On','Configured','Tested','Held Up'];
const VENUE_HEADERS = ['ID','VENUE','LOCATION','DECK','TV SIZE','QTY','AREA','SOURCE']
  .concat(VENUE_CHECKS, ['Status','Assigned To','Check By','Checked At','Remarks','Locked']);

function venueStatus_(read) {
  if (normalizeYesNo_(read('Configured')) === 'Yes') return 'Fully Configured';
  return VENUE_CHECKS.some(k => normalizeYesNo_(read(k)) === 'Yes') ? 'In Progress' : 'Not Started';
}
function venueSize_(v) { const m = String(v == null ? '' : v).match(/\d+/); return m ? m[0] : ''; }
/** "IPTV/GAMES" and "IPTV / GAMES" count as the same source. */
function venueSource_(v) { return String(v == null ? '' : v).replace(/\s*\/\s*/g, ' / ').replace(/\s+/g, ' ').trim().toUpperCase(); }
function venueArea_(v) { return /crew/i.test(String(v || '')) ? 'CREW' : 'PAX'; }

/** Creates an empty sheet (headers only) when missing, adds missing columns, fills blank IDs.
 *  Public areas are added with the web app upload (importVenueInventory) or typed into the sheet. */
function ensureVenueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(VENUE_SHEET);
  if (!sh || sh.getLastRow() < 1) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      sh = ss.getSheetByName(VENUE_SHEET) || ss.insertSheet(VENUE_SHEET);
      if (sh.getLastRow() < 1) {
        sh.getRange(1, 1, 1, VENUE_HEADERS.length).setValues([VENUE_HEADERS])
          .setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
        sh.setFrozenRows(1);
        bumpDataVersion_();
      }
    } finally { lock.releaseLock(); }
    return sh;
  }
  // Existing sheet: add any missing columns, give new rows an ID, fill blank checks with No.
  let lastCol = sh.getLastColumn();
  const m = {};
  sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].forEach((h, i) => { if (normalize_(h)) m[normalize_(h)] = i; });
  VENUE_HEADERS.forEach(h => {
    if (m[normalize_(h)] === undefined) {
      lastCol++;
      sh.getRange(1, lastCol).setValue(h).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
      m[normalize_(h)] = lastCol - 1;
    }
  });
  const last = sh.getLastRow();
  if (last < 2) return sh;
  const idCol = m[normalize_('ID')], vCol = m[normalize_('VENUE')];
  const data = sh.getRange(2, 1, last - 1, lastCol).getValues();
  let maxId = 0, idsChanged = false;
  data.forEach(r => { const n = String(r[idCol] || '').match(/^V(\d+)$/i); if (n) maxId = Math.max(maxId, Number(n[1])); });
  const seen = {};
  data.forEach(r => {
    if (!String(r[vCol] || '').trim() && !String(r[m[normalize_('LOCATION')]] || '').trim()) return;
    const id = String(r[idCol] || '').trim().toUpperCase();
    if (!id || seen[id]) { r[idCol] = 'V' + String(++maxId).padStart(3, '0'); idsChanged = true; }
    seen[String(r[idCol]).toUpperCase()] = true;
  });
  if (idsChanged) sh.getRange(2, idCol + 1, data.length, 1).setValues(data.map(r => [r[idCol]]));
  return sh;
}

const SCRIPT_VERSION = '2026-11-08 · Upload your own cabins and public areas';

function venueHeaderMap_(sh) {
  const m = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].forEach((h, i) => { if (normalize_(h)) m[normalize_(h)] = i; });
  return m;
}

/** Rewrites the Status column of the venue sheet for the given rows. */
function recalcVenueStatusRows_(sh, startRow, numRows) {
  const last = sh.getLastRow();
  numRows = Math.min(numRows, last - startRow + 1);
  if (startRow < 2 || numRows < 1) return;
  const m = venueHeaderMap_(sh), sCol = m[normalize_('Status')];
  if (sCol === undefined) return;
  const rows = sh.getRange(startRow, 1, numRows, sh.getLastColumn()).getValues();
  sh.getRange(startRow, sCol + 1, numRows, 1).setValues(rows.map(r => {
    if (!String(r[m[normalize_('VENUE')]] || '').trim() && !String(r[m[normalize_('LOCATION')]] || '').trim()) return [r[sCol]];
    return [venueStatus_(k => { const i = m[normalize_(k)]; return i === undefined ? 'No' : r[i]; })];
  }));
}

/** All public-area TV rows, straight from the sheet. */
function getVenueData_() {
  const sh = ensureVenueSheet_();
  if (sh.getLastRow() < 2) return [];
  const m = venueHeaderMap_(sh), cell = (r, k) => { const i = m[normalize_(k)]; return i === undefined || r[i] == null ? '' : r[i]; };
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().map((r, i) => {
    const venue = String(cell(r, 'VENUE')).trim(), location = String(cell(r, 'LOCATION')).trim();
    if (!venue && !location) return null;
    const deckRaw = cell(r, 'DECK');
    const o = {
      rowNumber: i + 2, id: String(cell(r, 'ID')).trim().toUpperCase(), venue, location,
      deck: deckRaw === '' ? '' : String(typeof deckRaw === 'number' ? Math.round(deckRaw) : deckRaw).trim(),
      size: venueSize_(cell(r, 'TV SIZE')), qty: Math.max(1, Math.round(Number(cell(r, 'QTY')) || 1)),
      area: venueArea_(cell(r, 'AREA')), source: venueSource_(cell(r, 'SOURCE')),
      assignedTo: techName_(cell(r, 'Assigned To')), checkBy: String(cell(r, 'Check By')),
      checkedAt: formatDate_(cell(r, 'Checked At')), remarks: String(cell(r, 'Remarks')), checks: {},
      locked: normalizeYesNo_(cell(r, 'Locked')) === 'Yes'
    };
    VENUE_CHECKS.forEach(k => o.checks[k] = normalizeYesNo_(cell(r, k)) === 'Yes');
    o.status = venueStatus_(k => o.checks[k] ? 'Yes' : 'No');
    return o;
  }).filter(Boolean);
}

/** TV counts (QTY-weighted) for the dashboard: whole ship, by size, by deck (with sources), by source, by area. */
function buildVenueStats_(venues) {
  const blank = () => { const b = { rows: 0, tvs: 0, fullyConfigured: 0, inProgress: 0, notStarted: 0 }; VENUE_CHECKS.forEach(k => b[k] = 0); return b; };
  const add = (b, v) => {
    b.rows++; b.tvs += v.qty;
    VENUE_CHECKS.forEach(k => { if (v.checks[k]) b[k] += v.qty; });
    if (v.status === 'Fully Configured') b.fullyConfigured += v.qty; else if (v.status === 'In Progress') b.inProgress += v.qty; else b.notStarted += v.qty;
  };
  const s = blank(); s.bySize = {}; s.byDeck = {}; s.bySource = {}; s.byArea = { PAX: blank(), CREW: blank() };
  venues.forEach(v => {
    add(s, v);
    add(s.byArea[v.area] || (s.byArea[v.area] = blank()), v);
    const z = v.size || '?'; add(s.bySize[z] || (s.bySize[z] = blank()), v);
    const d = v.deck || '?';
    if (!s.byDeck[d]) { s.byDeck[d] = blank(); s.byDeck[d].venues = {}; s.byDeck[d].sources = {}; }
    add(s.byDeck[d], v);
    s.byDeck[d].venues[v.venue] = true;
    if (v.source) s.byDeck[d].sources[v.source] = (s.byDeck[d].sources[v.source] || 0) + v.qty;
    const src = v.source || '(none)'; add(s.bySource[src] || (s.bySource[src] = blank()), v);
  });
  Object.keys(s.byDeck).forEach(d => { s.byDeck[d].venueCount = Object.keys(s.byDeck[d].venues).length; delete s.byDeck[d].venues; });
  return s;
}

/** Saves one public-area row. data = { id, checks:{label:bool}, baseChecks, remarks, baseRemarks, assignedTo, authToken, sessionId, requestId } */
function saveVenue(data) {
  if (!data || typeof data !== 'object') throw new Error('No venue information received.');
  const auth = authenticateToken_(data.authToken);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(data.sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  const requestId = String(data.requestId || '').trim(), cache = CacheService.getScriptCache();
  const cacheKey = requestId ? 'venue_save_' + requestId.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 80) : '';
  if (cacheKey) { const hit = cache.get(cacheKey); if (hit) { try { return JSON.parse(hit); } catch (_) {} } }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please retry.');
  try {
    const sh = ensureVenueSheet_(), m = venueHeaderMap_(sh), col = k => m[normalize_(k)], last = sh.getLastRow();
    const id = String(data.id || '').trim().toUpperCase();
    let rn = 0;
    if (id && last >= 2) {
      const ids = sh.getRange(2, col('ID') + 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim().toUpperCase() === id) { rn = i + 2; break; }
    }
    if (!rn) return { success: false, notFound: true, message: 'This TV location (' + id + ') is no longer in the Public Area TVs sheet. Refresh to load the current list.' };
    const row = sh.getRange(rn, 1, 1, sh.getLastColumn()).getValues()[0];
    const owner = techName_(row[col('Assigned To')]);
    if (owner && auth.username !== 'Admin' && owner !== auth.username) {
      return { success: false, forbidden: true, message: 'This location is assigned to ' + owner + '. Only ' + owner + ' or Admin can update it.' };
    }
    const checks = data.checks || {}, base = data.baseChecks || {}, changed = [], conflicts = [];
    VENUE_CHECKS.forEach(k => {
      if (!Object.prototype.hasOwnProperty.call(checks, k)) return;
      const sub = checks[k] ? 'Yes' : 'No', cur = normalizeYesNo_(row[col(k)]);
      const b = Object.prototype.hasOwnProperty.call(base, k) ? (base[k] ? 'Yes' : 'No') : cur;
      if (sub === b) return;
      if (cur !== b && cur !== sub) conflicts.push(k);
      else if (cur !== sub) changed.push({ field: k, value: sub });
    });
    if (Object.prototype.hasOwnProperty.call(data, 'remarks')) {
      const sub = String(data.remarks == null ? '' : data.remarks).slice(0, 2000), cur = String(row[col('Remarks')] == null ? '' : row[col('Remarks')]);
      const b = Object.prototype.hasOwnProperty.call(data, 'baseRemarks') ? String(data.baseRemarks == null ? '' : data.baseRemarks) : cur;
      if (sub !== b) { if (cur !== b && cur !== sub) conflicts.push('Remarks'); else if (cur !== sub) changed.push({ field: 'Remarks', value: sub }); }
    }
    if ((isChecklistLocked_() || normalizeYesNo_(col('Locked') === undefined ? '' : row[col('Locked')]) === 'Yes') && auth.username !== 'Admin') {
      const blocked = changed.filter(x => VENUE_CHECKS.includes(x.field) && x.value === 'No' && normalizeYesNo_(row[col(x.field)]) === 'Yes').map(x => x.field);
      if (blocked.length) return { success: false, locked: true, id,
        message: 'Admin has locked ticked items. You cannot untick: ' + blocked.join(', ') + '. Nothing was saved.' };
    }
    if (conflicts.length) return { success: false, conflict: true, message: 'Someone changed ' + conflicts.join(', ') + ' for ' + row[col('VENUE')] + '. Refresh and check the latest values before saving.' };
    if (Object.prototype.hasOwnProperty.call(data, 'assignedTo')) {
      const a = String(data.assignedTo || '').trim();
      if (a && !AUTHORIZED_TECHNICIANS.includes(a)) throw new Error('Choose a valid technician from the list.');
      if (a !== owner) {
        if (!canManageAssignments_(auth.username)) return { success: false, forbidden: true, message: 'Only Tech 3 or Admin can change who a location is assigned to.' };
        changed.push({ field: 'Assigned To', value: a });
      }
    }
    const now = new Date();
    changed.forEach(x => { row[col(x.field)] = x.value; });
    const status = venueStatus_(k => row[col(k)]);
    row[col('Status')] = status; row[col('Check By')] = auth.username; row[col('Checked At')] = now;
    sh.getRange(rn, 1, 1, row.length).setValues([row]);
    bumpDataVersion_();
    const where = String(row[col('VENUE')] || '') + (row[col('LOCATION')] ? ' — ' + row[col('LOCATION')] : '');
    const text = changed.map(x => x.field + ': ' + (x.value === '' ? '(cleared)' : x.value)).join('; ') || 'Save timestamp updated';
    appendAudit_(now, auth.username, 'VENUE', where + ' (' + id + ')', text, status, String(auth.sessionId || ''));
    queueTrackerEmail_('PUBLIC AREA TV UPDATED', { technician: auth.username, cabin: where, cabinType: 'VENUE', details: text, status, timestamp: formatDate_(now) });
    const out = { success: true, id, rowNumber: rn, status, checkBy: auth.username, checkedAt: formatDate_(now),
      assignedTo: techName_(row[col('Assigned To')]), remarks: String(row[col('Remarks')] == null ? '' : row[col('Remarks')]),
      dataVersion: getDataVersion(), message: where + ' saved.' };
    if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(out), 21600); } catch (_) {} }
    return out;
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

function setupVenueSheetFromMenu() {
  const existed = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VENUE_SHEET);
  const sh = ensureVenueSheet_();
  recalcVenueStatusRows_(sh, 2, Math.max(0, sh.getLastRow() - 1));
  bumpDataVersion_();
  SpreadsheetApp.getUi().alert((existed ? 'Public Area TVs sheet checked: ' : 'Public Area TVs sheet created: ') + Math.max(0, sh.getLastRow() - 1) + ' TV locations.' +
    (sh.getLastRow() < 2 ? ' Add public areas in the web app (Upload Cabins & Public Areas) or type them into the sheet.' : ''));
}

// ---------------------------------------------------------------------------
// LIVE SHEET SYNC
// ---------------------------------------------------------------------------
/** Simple trigger: runs automatically whenever someone edits the sheet by hand. */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (typeof FC_SHEET !== 'undefined' && sh.getName() === FC_SHEET) { fcOnEdit_(sh, e.range); return; }
    if (sh.getName() === VENUE_SHEET) {
      const vs = Math.max(2, e.range.getRow()), ve = e.range.getLastRow();
      if (ve >= vs) recalcVenueStatusRows_(sh, vs, ve - vs + 1);
      bumpDataVersion_();
      return;
    }
    if (sh.getName() !== SHEET_NAME) return;
    const start = Math.max(2, e.range.getRow()), end = e.range.getLastRow();
    if (end >= start) { syncCrewLinkedOnEdit_(sh, e.range); recalcStatusRows_(sh, start, end - start + 1); }
    bumpDataVersion_();
  } catch (err) { console.warn('onEdit sync failed: ' + err); }
}

/** Installable trigger: also catches inserted/deleted rows, sorting and pasted data. */
function handleTrackerChange(e) { bumpDataVersion_(); }

function installLiveSyncTriggers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'handleTrackerChange');
  if (!exists) ScriptApp.newTrigger('handleTrackerChange').forSpreadsheet(ss).onChange().create();
}

function enableLiveSyncFromMenu() { SpreadsheetApp.getUi().alert(setup()); }

function recalcAllStatusesFromMenu() {
  const sh = sheet_();
  if (sh.getLastRow() >= 2) recalcStatusRows_(sh, 2, sh.getLastRow() - 1);
  bumpDataVersion_();
  SpreadsheetApp.getUi().alert('All cabin statuses were recalculated.');
}

// ---------------------------------------------------------------------------
// SETUP — keeps the current sheet; an empty spreadsheet gets the tracker sheet with headers only
// ---------------------------------------------------------------------------
/** Writes the header row when the tracker sheet has none yet. */
function ensureCabinSheetHeaders_(sh) {
  const empty = sh.getLastRow() < 1 || !sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0].some(v => String(v).trim());
  if (!empty) return;
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this Apps Script project from the Google Spreadsheet you want to use.');
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  ensureCabinSheetHeaders_(sh);
  ensureTrackerColumns_(sh);
  if (sh.getLastRow() >= 2) recalcStatusRows_(sh, 2, sh.getLastRow() - 1);
  ensureVenueSheet_();
  ensureActivitySheets_();
  ensureEmailQueueSheet_();
  installTrackerEmailQueueTrigger_();
  installLiveSyncTriggers_();
  try { installMaintenanceTriggers_(); PropertiesService.getScriptProperties().setProperty(MAINT_TRIGGERS_PROPERTY, new Date().toISOString()); } catch (err) { console.warn('Maintenance triggers not installed: ' + err); }
  bumpDataVersion_();
  SpreadsheetApp.flush();
  const rows = Math.max(0, sh.getLastRow() - 1);
  return rows
    ? 'Setup checked: your current cabin list was kept (' + rows + ' rows). Live sheet sync is on.'
    : 'Setup done. The cabin list is empty: sign in to the web app as the admin and use "Upload Cabins & Public Areas" ' +
      'to add your cabins, crew offices and public areas. Live sheet sync is on.';
}

// ---------------------------------------------------------------------------
// UPLOAD YOUR OWN CABINS, CREW OFFICES AND PUBLIC AREAS (admin, web app)
// The browser reads the .xlsx / .csv file and sends plain rows; these functions write them to the sheets.
// mode 'merge'   = add new rows, update details of rows already there, keep everything else.
// mode 'replace' = the sheet ends up with exactly the rows in the file; rows in both keep their ticks / remarks.
// ---------------------------------------------------------------------------
const INVENTORY_MAX_ROWS = 10000;

function invText_(v, max) {
  if (typeof v === 'number') v = Number.isInteger(v) ? String(v) : String(v);
  return String(v == null ? '' : v).replace(/ /g, ' ').replace(/\s+/g, ' ').trim().replace(/\.0+$/, '').slice(0, max || 100);
}
/** PAX / CREW from what the file says ("Guest", "Crew office", …); blank falls back to the cabin number. */
function invType_(type, cabin) {
  const t = normalize_(type);
  if (/crew|office|officer|staff/.test(t)) return 'CREW';
  if (/pax|guest|passenger/.test(t)) return 'PAX';
  return inferType_('', cabin);
}
/** Cabin number as stored: upper case, no spaces; a crew number starting with a digit gets "C" (1234 -> C1234). */
function invCabin_(type, v) {
  let s = invText_(v, 40).toUpperCase().replace(/\s+/g, '');
  if (s && type === 'CREW' && /^\d/.test(s)) s = 'C' + s;
  return s;
}
function invCheckColumns_() {
  const set = {};
  ALL_TV_FIELDS.concat(CABIN_EXTRA_FIELDS, COMMON_CHECK_FIELDS, ['Locked']).forEach(f => set[normalize_(f)] = true);
  return set;
}

/** items: [{ type, cabin, deck, fz, tvCount }] */
function importCabinInventory(token, sessionId, items, mode) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'The file has no cabins.' };
  if (items.length > INVENTORY_MAX_ROWS) return { success: false, message: 'Too many rows in one upload (maximum ' + INVENTORY_MAX_ROWS + ').' };
  const replace = mode === 'replace';

  const list = [], seen = {};
  let dup = 0, bad = 0;
  items.forEach(it => {
    if (!it || typeof it !== 'object') { bad++; return; }
    const type = invType_(it.type, invText_(it.cabin, 40));
    const cabin = invCabin_(type, it.cabin);
    if (!cabin) { bad++; return; }
    const key = type + ':' + cabin;
    if (seen[key]) { dup++; return; }
    seen[key] = true;
    list.push({ type, cabin, key, deck: invText_(it.deck, 20), fz: invText_(it.fz, 20), tvCount: parseTvCount_(it.tvCount) });
  });
  if (!list.length) return { success: false, message: 'No cabin numbers found in the file.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = sheet_();
    ensureCabinSheetHeaders_(sh);
    const m = headerMap_(sh), width = sh.getLastColumn(), last = sh.getLastRow();
    const col = h => m[normalize_(h)];
    const ci = col('CABIN NO'), ti = col('CABIN TYPE'), di = col('DECK'), fi = col('FZ'), tc = col('TV Count'), si = col('Status');
    const existing = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    const byKey = {};
    existing.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c && byKey[cabinKey_(r[ti], c)] === undefined) byKey[cabinKey_(r[ti], c)] = i; });

    const checks = invCheckColumns_(), headers = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
    const newRow = x => {
      const r = headers.map(h => checks[normalize_(h)] ? 'No' : '');
      r[ti] = x.type; r[ci] = x.cabin; r[di] = x.deck; r[fi] = x.fz; r[tc] = x.tvCount || 1; r[si] = 'Not Started';
      return r;
    };

    const used = {}, keep = [], add = [];
    let updated = 0;
    const changedCols = {};
    list.forEach(x => {
      const i = byKey[x.key];
      if (i === undefined) { add.push(newRow(x)); return; }
      used[i] = true;
      const r = existing[i];
      let changed = false;
      const set = (c, v) => { if (c !== undefined && v !== '' && String(r[c] == null ? '' : r[c]).trim() !== String(v)) { r[c] = v; changedCols[c] = true; changed = true; } };
      set(ti, x.type); set(di, x.deck); set(fi, x.fz);
      if (x.tvCount && parseTvCount_(r[tc]) !== x.tvCount) { r[tc] = x.tvCount; changedCols[tc] = true; changed = true; }
      if (changed) updated++;
      keep.push(r);
    });
    const removedRows = existing.filter((r, i) => !used[i] && String(r[ci] == null ? '' : r[ci]).trim());

    if (replace) {
      const out = keep.concat(add);
      if (last >= 2) sh.getRange(2, 1, last - 1, width).clearContent();
      if (out.length) sh.getRange(2, 1, out.length, width).setValues(out);
      if (out.length) recalcStatusRows_(sh, 2, out.length);
    } else {
      // Only the changed columns of existing rows are rewritten, so formulas and formatting elsewhere stay.
      Object.keys(changedCols).map(Number).forEach(c => sh.getRange(2, c + 1, existing.length, 1).setValues(existing.map(r => [r[c]])));
      if (add.length) sh.getRange(last + 1, 1, add.length, width).setValues(add);
    }
    const ca = col('Checked At'), rows = Math.max(1, sh.getLastRow() - 1);
    if (ca !== undefined) sh.getRange(2, ca + 1, rows, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    const removed = replace ? removedRows.length : 0;
    const summary = 'Cabin upload (' + (replace ? 'replace list' : 'add / update') + '): ' + add.length + ' added, ' + updated + ' updated' +
      (replace ? ', ' + removed + ' removed' : '') + '. Total now ' + Math.max(0, sh.getLastRow() - 1) + '.';
    try { appendAudit_(new Date(), a.auth.username, 'CABINS', '', summary, 'Upload', String(a.auth.sessionId || '')); } catch (_) {}
    bumpDataVersion_();
    SpreadsheetApp.flush();
    const notes = [];
    if (dup) notes.push(dup + ' duplicate row(s) in the file skipped');
    if (bad) notes.push(bad + ' row(s) without a cabin number skipped');
    return { success: true, added: add.length, updated, removed, dataVersion: getDataVersion(),
      message: summary + (notes.length ? ' ' + notes.join('; ') + '.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

/** items: [{ venue, location, deck, size, qty, area, source }] */
function importVenueInventory(token, sessionId, items, mode) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'The file has no public areas.' };
  if (items.length > INVENTORY_MAX_ROWS) return { success: false, message: 'Too many rows in one upload (maximum ' + INVENTORY_MAX_ROWS + ').' };
  const replace = mode === 'replace';
  const keyOf = (venue, location, deck, size) => [normalize_(venue), normalize_(location), invText_(deck, 20), venueSize_(size)].join('|');

  const list = [], seen = {};
  let dup = 0, bad = 0;
  items.forEach(it => {
    if (!it || typeof it !== 'object') { bad++; return; }
    const venue = invText_(it.venue, 120).toUpperCase(), location = invText_(it.location, 160).toUpperCase();
    if (!venue && !location) { bad++; return; }
    const deck = invText_(it.deck, 20), size = venueSize_(it.size);
    const key = keyOf(venue, location, deck, size);
    if (seen[key]) { dup++; return; }
    seen[key] = true;
    const qty = Math.max(1, Math.min(999, Math.round(Number(String(it.qty == null ? '' : it.qty).replace(/[^\d.]/g, '')) || 1)));
    const areaText = invText_(it.area, 40);
    list.push({ key, venue, location, deck: /^\d+$/.test(deck) ? Number(deck) : deck, size: size ? size + '"' : '', qty,
      area: areaText ? (venueArea_(areaText) === 'CREW' ? 'CREW AREA' : 'PAX AREA') : '', source: venueSource_(it.source) });
  });
  if (!list.length) return { success: false, message: 'No venues or locations found in the file.' };

  const sh = ensureVenueSheet_();   // outside the lock: it may take the lock itself to create the sheet
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const m = venueHeaderMap_(sh), width = sh.getLastColumn(), last = sh.getLastRow();
    const col = h => m[normalize_(h)];
    const existing = last >= 2 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    const cell = (r, h) => { const i = col(h); return i === undefined || r[i] == null ? '' : r[i]; };
    const byKey = {};
    let maxId = 0;
    existing.forEach((r, i) => {
      const n = String(cell(r, 'ID')).match(/^V(\d+)$/i); if (n) maxId = Math.max(maxId, Number(n[1]));
      if (!String(cell(r, 'VENUE')).trim() && !String(cell(r, 'LOCATION')).trim()) return;
      const k = keyOf(cell(r, 'VENUE'), cell(r, 'LOCATION'), cell(r, 'DECK'), cell(r, 'TV SIZE'));
      if (byKey[k] === undefined) byKey[k] = i;
    });
    const headers = sh.getRange(1, 1, 1, width).getDisplayValues()[0];
    const checkSet = {}; VENUE_CHECKS.concat(['Locked']).forEach(k => checkSet[normalize_(k)] = true);

    const used = {}, keep = [], add = [], changedCols = {};
    let updated = 0;
    list.forEach(x => {
      const i = byKey[x.key];
      if (i === undefined) {
        const r = headers.map(h => checkSet[normalize_(h)] ? 'No' : '');
        const put = (h, v) => { const c = col(h); if (c !== undefined) r[c] = v; };
        put('ID', 'V' + String(++maxId).padStart(3, '0')); put('VENUE', x.venue); put('LOCATION', x.location); put('DECK', x.deck);
        put('TV SIZE', x.size); put('QTY', x.qty); put('AREA', x.area || 'PAX AREA'); put('SOURCE', x.source); put('Status', 'Not Started');
        add.push(r);
        return;
      }
      used[i] = true;
      const r = existing[i];
      let changed = false;
      const set = (h, v) => { const c = col(h); if (c !== undefined && v !== '' && String(r[c] == null ? '' : r[c]).trim() !== String(v)) { r[c] = v; changedCols[c] = true; changed = true; } };
      set('QTY', x.qty); set('AREA', x.area); set('SOURCE', x.source);
      if (changed) updated++;
      keep.push(r);
    });
    const removed = replace ? existing.filter((r, i) => !used[i] && (String(cell(r, 'VENUE')).trim() || String(cell(r, 'LOCATION')).trim())).length : 0;

    if (replace) {
      const out = keep.concat(add);
      if (last >= 2) sh.getRange(2, 1, last - 1, width).clearContent();
      if (out.length) { sh.getRange(2, 1, out.length, width).setValues(out); recalcVenueStatusRows_(sh, 2, out.length); }
    } else {
      Object.keys(changedCols).map(Number).forEach(c => sh.getRange(2, c + 1, existing.length, 1).setValues(existing.map(r => [r[c]])));
      if (add.length) sh.getRange(last + 1, 1, add.length, width).setValues(add);
    }
    const ca = col('Checked At');
    if (ca !== undefined) sh.getRange(2, ca + 1, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    const summary = 'Public area upload (' + (replace ? 'replace list' : 'add / update') + '): ' + add.length + ' added, ' + updated + ' updated' +
      (replace ? ', ' + removed + ' removed' : '') + '. Total now ' + Math.max(0, sh.getLastRow() - 1) + ' TV locations.';
    try { appendAudit_(new Date(), a.auth.username, 'VENUE', '', summary, 'Upload', String(a.auth.sessionId || '')); } catch (_) {}
    bumpDataVersion_();
    SpreadsheetApp.flush();
    const notes = [];
    if (dup) notes.push(dup + ' duplicate row(s) in the file skipped');
    if (bad) notes.push(bad + ' row(s) without a venue or location skipped');
    return { success: true, added: add.length, updated, removed, dataVersion: getDataVersion(),
      message: summary + (notes.length ? ' ' + notes.join('; ') + '.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

// ---------------------------------------------------------------------------
// DATA READ — always straight from the sheet
// ---------------------------------------------------------------------------
function getData_() {
  let sh = sheet_();
  if (sh.getLastRow() < 1) setup();   // brand-new spreadsheet: headers + triggers
  ensureCrewLanSynced_();
  sh = sheet_();
  const map = headerMap_(sh);
  if (sh.getLastRow() < 2) return [];
  const col = n => map[normalize_(n)];
  const cell = (r, n) => { const i = col(n); return i === undefined || r[i] == null ? '' : r[i]; };
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues().map((r, i) => {
    const cabin = String(cell(r, 'CABIN NO')).trim();
    if (!cabin) return null;
    const cabinType = inferType_(cell(r, 'CABIN TYPE'), cabin);
    const o = {
      rowNumber: i + 2, cabinType, cabin,
      deck: String(cell(r, 'DECK')).trim(), fz: String(cell(r, 'FZ')).trim(),
      tvCount: parseTvCount_(cell(r, 'TV Count')) || getTvCount_(cabinType, cabin),
      checkBy: String(cell(r, 'Check By')), assignedTo: techName_(cell(r, 'Assigned To')),
      status: '', checkedAt: formatDate_(cell(r, 'Checked At')), comments: String(cell(r, 'Comments')),
      whatToCheck: String(cell(r, 'What to Check')),
      locked: normalizeYesNo_(cell(r, 'Locked')) === 'Yes'   // set by Admin (per-cabin lock)
    };
    ALL_TV_FIELDS.concat(COMMON_CHECK_FIELDS, CABIN_EXTRA_FIELDS).forEach(f => {
      const idx = col(f);
      o[f] = idx !== undefined && normalizeYesNo_(r[idx]) === 'Yes';   // works for Yes/No text and TRUE/FALSE checkboxes
    });
    if (cabinType !== 'PAX') o[SECURITY_DOOR] = false; // PAX only
    o.noTvHere = !!o['No TV Here'];
    o.status = calculateCabinStatusFromFields_(f => o[f] ? 'Yes' : 'No', o.tvCount, cabinType);
    return o;
  }).filter(Boolean);
}

/** Authenticated cabin list: Admin sees all, others only their assigned cabins. */
function getData(token) {
  const auth = authenticateToken_(token);
  if (!auth) return [];
  const all = getData_();
  return auth.username === 'Admin' ? all : all.filter(c => c.assignedTo === auth.username);
}

// ---------------------------------------------------------------------------
// EMAIL NOTIFICATIONS
// ---------------------------------------------------------------------------
function configureTrackerEmailNotifications() {
  const ui = SpreadsheetApp.getUi();
  const current = PropertiesService.getScriptProperties().getProperty(TRACKER_EMAIL_PROPERTY) || '';
  const response = ui.prompt('Tracker Email Notifications',
    'Enter the email address that should receive tracker alerts.' + (current ? '\nCurrent recipient: ' + current : ''),
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return 'No changes made.';
  const email = String(response.getResponseText() || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Please enter a valid email address.');
  PropertiesService.getScriptProperties().setProperty(TRACKER_EMAIL_PROPERTY, email);
  try { installTrackerEmailQueueTrigger_(); } catch (err) { console.warn('Email queue trigger could not be installed: ' + err); }
  try {
    const quota = MailApp.getRemainingDailyQuota();
    if (quota <= 0) return 'Recipient saved: ' + email + '. Daily email quota is exhausted; notifications resume when it resets.';
    return 'Recipient saved: ' + email + '. Remaining daily email quota: ' + quota + '.';
  } catch (err) {
    return 'Recipient saved: ' + email + '. Quota could not be checked: ' + String(err && err.message ? err.message : err);
  }
}

function disableTrackerEmailNotifications() {
  PropertiesService.getScriptProperties().deleteProperty(TRACKER_EMAIL_PROPERTY);
  return 'Tracker email notifications have been disabled.';
}

function testTrackerEmailNotification() {
  const email = PropertiesService.getScriptProperties().getProperty(TRACKER_EMAIL_PROPERTY);
  if (!email) throw new Error('Configure the notification recipient first by running configureTrackerEmailNotifications().');
  try {
    if (MailApp.getRemainingDailyQuota() <= 0) return 'Test email not sent: daily email quota is exhausted.';
  } catch (err) {
    return 'Test email not sent because the quota could not be checked: ' + String(err && err.message ? err.message : err);
  }
  const result = sendTrackerEmail_('TEST', { technician: 'Tracker administrator', details: 'This is a test notification from the Cabin TV Installation Tracker.', timestamp: formatDate_(new Date()) });
  return result.sent ? 'Test email sent to ' + email + '.' : 'Test email not sent: ' + result.reason;
}

function trackerEmailQuotaDay_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

function sendTrackerEmail_(eventName, info) {
  try {
    const props = PropertiesService.getScriptProperties();
    const recipient = props.getProperty(TRACKER_EMAIL_PROPERTY);
    if (!recipient) return { sent: false, reason: 'No notification recipient configured.' };
    if (props.getProperty('TRACKER_EMAIL_QUOTA_EXHAUSTED_DAY') === trackerEmailQuotaDay_()) {
      return { sent: false, reason: 'Daily email quota already exhausted; retry after the quota resets.' };
    }
    if (MailApp.getRemainingDailyQuota() <= 0) {
      props.setProperty('TRACKER_EMAIL_QUOTA_EXHAUSTED_DAY', trackerEmailQuotaDay_());
      return { sent: false, reason: 'Google Apps Script daily email quota is exhausted.' };
    }
    const d = info || {};
    const lines = [
      'Cabin TV Installation Tracker', '',
      'Event: ' + String(eventName || 'Activity'),
      'Technician: ' + String(d.technician || 'Unknown'),
      d.cabin ? 'Cabin: ' + String(d.cabinType || '') + ' ' + String(d.cabin) : '',
      d.details ? 'Details: ' + String(d.details) : '',
      d.status ? 'Status: ' + String(d.status) : '',
      'Time: ' + (d.timestamp || formatDate_(new Date())), '',
      'This is an automated notification from the tracker.'
    ].filter(Boolean);
    MailApp.sendEmail({ to: recipient, subject: '[TV Tracker] ' + String(eventName || 'Activity'), body: lines.join('\n'), name: 'Cabin TV Tracker' });
    return { sent: true };
  } catch (err) {
    const message = String(err && err.message ? err.message : err);
    if (/too many times for one day|daily.*quota|quota.*exhaust/i.test(message)) {
      PropertiesService.getScriptProperties().setProperty('TRACKER_EMAIL_QUOTA_EXHAUSTED_DAY', trackerEmailQuotaDay_());
    }
    console.error('Tracker notification email failed: ' + message);
    return { sent: false, reason: message };
  }
}

function ensureEmailQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(EMAIL_QUEUE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EMAIL_QUEUE_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['Queued At','Event','Technician','Cabin Type','Cabin','Details','Status','Timestamp','Processed']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
  }
  return sh;
}

function queueTrackerEmail_(eventName, info) {
  try {
    if (!PropertiesService.getScriptProperties().getProperty(TRACKER_EMAIL_PROPERTY)) return { queued: false, reason: 'No recipient configured.' };
    const d = info || {};
    ensureEmailQueueSheet_().appendRow([new Date(), String(eventName || 'Activity'), String(d.technician || ''), String(d.cabinType || ''),
      String(d.cabin || ''), String(d.details || ''), String(d.status || ''), String(d.timestamp || formatDate_(new Date())), '']);
    return { queued: true };
  } catch (err) {
    console.error('Could not queue tracker email: ' + err);
    return { queued: false, reason: String(err) };
  }
}

function installTrackerEmailQueueTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processTrackerEmailQueue');
  if (!exists) ScriptApp.newTrigger('processTrackerEmailQueue').timeBased().everyMinutes(1).create();
}

function processTrackerEmailQueue() {
  if (!PropertiesService.getScriptProperties().getProperty(TRACKER_EMAIL_PROPERTY)) return; // no recipient: nothing to send
  const sh = ensureEmailQueueSheet_(), last = sh.getLastRow();
  if (last < 2) return;
  const rows = sh.getRange(2, 1, last - 1, 9).getValues();
  let processed = 0;
  for (let i = 0; i < rows.length && processed < 20; i++) {
    if (rows[i][8]) continue;
    const r = rows[i];
    const result = sendTrackerEmail_(String(r[1] || 'Activity'), { technician: r[2], cabinType: r[3], cabin: r[4], details: r[5], status: r[6], timestamp: r[7] });
    if (result && result.sent) { sh.getRange(i + 2, 9).setValue(new Date()); processed++; }
  }
}

// ---------------------------------------------------------------------------
// INDIVIDUAL TECHNICIAN AUTHENTICATION
// ---------------------------------------------------------------------------
function accountName_(username) {
  const n = String(username || '').trim().toLowerCase();
  return TECHNICIANS.find(t => t.toLowerCase() === n.replace(/\s+/g, ' ')) || '';
}
function accountKey_(username) {
  const name = accountName_(username);
  return name ? name.toUpperCase().replace(/\s+/g, '_') : ''; // TECH_1, ADMIN, … TV_CHECK_1
}
function constantTimeEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  let diff = a.length ^ b.length;
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function passwordHash_(password, salt) {
  let v = String(salt) + '|' + String(password);
  for (let i = 0; i < 5000; i++) v = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8));
  return v;
}

/** Quicker salted hash (works on bytes, fewer rounds). Used after the first sign-in; see loginTechnician. */
function passwordHashFast_(password, salt) {
  let bytes = Utilities.newBlob(String(salt) + '|' + String(password)).getBytes();
  for (let i = 0; i < 600; i++) bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  return Utilities.base64EncodeWebSafe(bytes);
}
/** Saves the quick hash, tied to the current main hash: if the password is reset, the quick hash no longer matches and is ignored. */
function storeFastHash_(key, password, mainHash) {
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('TRACKER_USER_' + key + '_FAST',
    JSON.stringify({ of: String(mainHash).slice(-16), salt: salt, hash: passwordHashFast_(password, salt) }));
}
function readFastHash_(key, mainHash) {
  try {
    const f = JSON.parse(PropertiesService.getScriptProperties().getProperty('TRACKER_USER_' + key + '_FAST') || 'null');
    return f && f.salt && f.hash && f.of === String(mainHash).slice(-16) ? f : null;
  } catch (_) { return null; }
}

function saveTechnicianPasswordSetup(adminKey, username, password) {
  const props = PropertiesService.getScriptProperties();
  const expected = props.getProperty('TRACKER_ADMIN_SETUP_KEY');
  if (!expected) throw new Error('Admin setup key is not configured in Script Properties. Add TRACKER_ADMIN_SETUP_KEY first.');
  if (!constantTimeEqual_(String(adminKey || ''), expected)) throw new Error('Incorrect admin setup key.');
  const name = accountName_(username), key = accountKey_(username), pw = String(password || '');
  if (!name || !key) throw new Error('Only these accounts can have editor passwords: ' + TECHNICIANS.join(', ') + '.');
  if (pw.length < 12) throw new Error('Password must be at least 12 characters.');
  const salt = Utilities.getUuid() + '-' + Utilities.getUuid();
  const mainHash = passwordHash_(pw, salt);
  props.setProperties({ ['TRACKER_USER_' + key + '_SALT']: salt, ['TRACKER_USER_' + key + '_HASH']: mainHash });
  storeFastHash_(key, pw, mainHash);
  return { success: true, message: 'Password saved for ' + name + '. The password is stored as a salted hash.' };
}

/** Strong random password, e.g. "TVC3-k7Pq-9mXa-3RwH" (no look-alike characters). */
function randomPassword_(prefix) {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid() + Date.now());
  let s = '';
  for (let i = 0; i < 12; i++) s += chars[(bytes[i] + 256) % chars.length];
  return prefix + '-' + s.slice(0, 4) + '-' + s.slice(4, 8) + '-' + s.slice(8, 12);
}

/** Menu: creates passwords for TV Check accounts and shows them once. Only salted hashes are stored. */
function createTvCheckPasswordsFromMenu() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const hasPw = n => !!props.getProperty('TRACKER_USER_' + accountKey_(n) + '_HASH');
  const without = TV_CHECK_ACCOUNTS.filter(n => !hasPw(n));
  const r = ui.prompt('TV Check passwords',
    'Which accounts? Type the numbers, e.g.  1-6  or  2,4,5,6  or  4  or  ALL\n' +
    'Leave empty for the accounts that have no password yet (' + (without.length ? without.join(', ') : 'none') + ').\n' +
    'An account that already has a password gets a NEW one (the old one stops working).', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const typed = r.getResponseText().trim();
  // Understands "1,2,5", "4", "1-6", "2 to 4", "tv check 2, tv check 5" and "all".
  const nums = new Set();
  if (/\ball\b/i.test(typed)) TV_CHECK_ACCOUNTS.forEach((_, i) => nums.add(i + 1));
  typed.replace(/(\d+)\s*(?:-|–|—|to)\s*(\d+)|(\d+)/gi, (_, a, b, one) => {
    if (one) nums.add(Number(one));
    else for (let n = Math.min(+a, +b); n <= Math.max(+a, +b); n++) nums.add(n);
    return '';
  });
  const list = typed
    ? [...nums].sort((x, y) => x - y).filter(n => n >= 1 && n <= TV_CHECK_ACCOUNTS.length).map(n => 'TV Check ' + n)
    : without;
  if (!list.length) { ui.alert(typed ? 'No valid account numbers (use 1–' + TV_CHECK_ACCOUNTS.length + ').' : 'All TV Check accounts already have a password. Type the numbers to reset them.'); return; }
  const lines = list.map(name => {
    const pw = randomPassword_('TVC' + name.replace(/\D/g, ''));
    const salt = Utilities.getUuid() + '-' + Utilities.getUuid(), key = accountKey_(name);
    const mainHash = passwordHash_(pw, salt);
    props.setProperties({ ['TRACKER_USER_' + key + '_SALT']: salt, ['TRACKER_USER_' + key + '_HASH']: mainHash });
    storeFastHash_(key, pw, mainHash);
    return name + ':   ' + pw;
  });
  try { appendAudit_(new Date(), 'Admin (spreadsheet)', 'SESSION', '', 'Passwords set for ' + list.join(', '), 'Settings', ''); } catch (_) {}
  ui.alert('TV Check passwords — write these down NOW',
    lines.join('\n') + '\n\nThey are not stored and cannot be shown again (only a secure hash is kept). ' +
    'To change one later, run this menu again with its number.', ui.ButtonSet.OK);
}

function authCache_() { return CacheService.getScriptCache(); }
function authCacheKey_(token) {
  return 'TRACKER_AUTH_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token || ''))).slice(0, 40);
}

/** Sessions live in Script Properties (CacheService may evict early). Sliding 6-hour inactivity timeout. */
function authenticateToken_(token) {
  token = String(token || '').trim();
  if (!token) return null;
  const props = PropertiesService.getScriptProperties(), key = authCacheKey_(token), raw = props.getProperty(key);
  if (!raw) return null;
  try {
    const auth = JSON.parse(raw);
    if (!accountName_(auth.username) || !auth.sessionId || !auth.createdAt) return null;
    if (Date.now() - Number(auth.lastSeenAt || auth.createdAt) > AUTH_CACHE_TTL_SECONDS * 1000) { props.deleteProperty(key); return null; }
    // Only rewrite the session every 5 minutes: writing Script Properties on every call slowed every request.
    if (Date.now() - Number(auth.lastSeenAt || 0) > 300000) { auth.lastSeenAt = Date.now(); props.setProperty(key, JSON.stringify(auth)); }
    return auth;
  } catch (_) { return null; }
}

function verifyTechnicianSession(token) {
  const a = authenticateToken_(token);
  return a ? { success: true, username: a.username, expiresInSeconds: AUTH_CACHE_TTL_SECONDS }
           : { success: false, message: 'Your editing session expired. Please sign in again.' };
}

function loginTechnician(username, password, sessionId) {
  const name = accountName_(username), key = accountKey_(username), pw = String(password || ''), sid = String(sessionId || '').trim();
  if (!name || !key) return { success: false, message: 'Choose an editor account: ' + TECHNICIANS.join(', ') + '.' };
  if (!sid) return { success: false, message: 'Browser session is missing. Reload the tracker and try again.' };
  const cache = authCache_(), failKey = 'TRACKER_LOGIN_FAIL_' + key, failed = Number(cache.get(failKey) || 0);
  if (failed >= 5) return { success: false, message: 'Too many unsuccessful attempts. Wait 15 minutes before trying again.' };
  const props = PropertiesService.getScriptProperties();
  if (!needsPassword_(name)) {   // technician accounts: no password
    const now = Date.now(), token = Utilities.getUuid() + '-' + Utilities.getUuid();
    props.setProperty(authCacheKey_(token), JSON.stringify({ username: name, sessionId: sid, createdAt: now, lastSeenAt: now }));
    return { success: true, token: token, username: name };
  }
  const salt = props.getProperty('TRACKER_USER_' + key + '_SALT'), expected = props.getProperty('TRACKER_USER_' + key + '_HASH');
  if (!salt || !expected) return { success: false, message: 'Password has not been configured for ' + name + '. Ask the tracker administrator to set it.' };
  const fast = readFastHash_(key, expected);
  const ok = fast ? constantTimeEqual_(passwordHashFast_(pw, fast.salt), fast.hash)
                  : constantTimeEqual_(passwordHash_(pw, salt), expected);
  if (!ok) { cache.put(failKey, String(failed + 1), 900); return { success: false, message: 'Incorrect username or password.' }; }
  if (!fast) { try { storeFastHash_(key, pw, expected); } catch (_) {} } // next sign-in uses the quick check
  cache.remove(failKey);
  const now = Date.now(), token = Utilities.getUuid() + '-' + Utilities.getUuid();
  props.setProperty(authCacheKey_(token), JSON.stringify({ username: name, sessionId: sid, createdAt: now, lastSeenAt: now }));
  return { success: true, token: token, username: name };
}

function registerTechnicianLogin(token) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Session expired; activity registration was skipped.' };
  const now = new Date(), lock = LockService.getScriptLock();
  if (!lock.tryLock(1200)) return { success: true, transient: true };
  try {
    upsertActiveUser_(auth.sessionId, auth.username, '', now);
  } finally { if (lock.hasLock()) lock.releaseLock(); }
  appendAudit_(now, auth.username, 'SESSION', '', 'Signed in with individual account', 'Session Started', auth.sessionId);
  return { success: true };
}

function notifyTechnicianLogin(token) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, message: 'Session expired; login notification was skipped.' };
  // Queued: the every-minute job sends it, so signing in does not wait for the email.
  const result = queueTrackerEmail_('TECHNICIAN SIGNED IN', { technician: auth.username, details: auth.username + ' successfully signed in.', timestamp: formatDate_(new Date()) });
  return { success: !!result.queued, message: result.reason || '' };
}

function upsertActiveUser_(sessionId, username, cabin, when) {
  const sh = ensureActivitySheets_().active, last = sh.getLastRow();
  let row = -1, work = 'Working';
  if (last >= 2) {
    const found = sh.getRange(2, 1, last - 1, 1).createTextFinder(sessionId).matchEntireCell(true).findNext();
    if (found) { row = found.getRow(); work = String(sh.getRange(row, 6).getValue() || 'Working'); }
  }
  const rec = [sessionId, username, cabin, when, 'Active', work];
  if (row < 0) sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([rec]);
  else sh.getRange(row, 1, 1, 6).setValues([rec]);
}

function heartbeatUser(sessionId, currentCabin, token) {
  const auth = authenticateToken_(token);
  sessionId = String(sessionId || '').trim();
  if (!auth || auth.sessionId !== sessionId) return { success: false, message: 'Editing session expired. Please sign in again.' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(700)) return { success: true, transient: true, technician: auth.username };
  try { upsertActiveUser_(sessionId, auth.username, String(currentCabin || '').trim(), new Date()); }
  finally { if (lock.hasLock()) lock.releaseLock(); }
  return { success: true, technician: auth.username, lastSeen: formatDate_(new Date()) };
}

function logoutTechnician(token, sessionId) {
  const auth = authenticateToken_(token);
  sessionId = String(sessionId || '').trim();
  if (auth && auth.sessionId === sessionId) endUserSession(sessionId, token);
  if (token) PropertiesService.getScriptProperties().deleteProperty(authCacheKey_(token));
  return { success: true };
}

// ---------------------------------------------------------------------------
// SAVE A CABIN — located by cabin type + number, so row changes in the sheet are safe
// ---------------------------------------------------------------------------
function saveTV(rowNumber, data) {
  if (!data || typeof data !== 'object') throw new Error('No cabin information received.');
  const auth = authenticateToken_(data.authToken);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(data.sessionId || '').trim() !== String(auth.sessionId || '')) {
    return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  }

  // Idempotency: a retried request returns the original result.
  const requestId = String(data.requestId || '').trim();
  const cache = CacheService.getScriptCache();
  const cacheKey = requestId ? 'paxcrew_save_' + requestId.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 80) : '';
  if (cacheKey) { const hit = cache.get(cacheKey); if (hit) { try { return JSON.parse(hit); } catch (_) {} } }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please retry this cabin.');
  try {
    const sh = sheet_(), m = headerMap_(sh), col = n => m[normalize_(n)];
    const rn = resolveCabinRow_(sh, m, data.cabinType, data.cabin, Number(rowNumber));
    if (!rn) {
      return { success: false, notFound: true, cabin: String(data.cabin || ''), cabinType: String(data.cabinType || ''),
        message: 'Cabin ' + String(data.cabin || '') + ' is no longer in the tracker sheet (it was renamed, retyped or removed). Refresh to load the current list.' };
    }
    const row = sh.getRange(rn, 1, 1, sh.getLastColumn()).getValues()[0];
    const cabin = String(row[col('CABIN NO')] == null ? '' : row[col('CABIN NO')]).trim();
    const type = inferType_(row[col('CABIN TYPE')], cabin);
    const isPax = type === 'PAX';
    if (!cabin) throw new Error('Cabin record was not found. Refresh and try again.');

    const assignedOwner = techName_(row[col('Assigned To')]);
    if (auth.username !== 'Admin' && assignedOwner !== auth.username) {
      return { success: false, forbidden: true, message: 'You can only view and update cabins assigned to your account.' };
    }

    const currentStamp = formatDate_(row[col('Checked At')]);
    const baseStamp = String(data.baseCheckedAt || '');
    const baseChecks = data.baseChecks && typeof data.baseChecks === 'object' ? data.baseChecks : null;
    const findKey = (obj, f) => obj ? Object.keys(obj).find(k => normalize_(k) === normalize_(f)) : undefined;
    const changed = [], conflicts = [];

    ALL_TV_FIELDS.concat(COMMON_CHECK_FIELDS, CABIN_EXTRA_FIELDS).forEach(f => {
      // AP and Security to Open the Door only exist on PAX cabins. (Patch cable is "LAN Connected" on CREW cabins.)
      if (!isPax && ['AP', SECURITY_DOOR].includes(f.replace(/^TV [23] /, ''))) return;
      const key = findKey(data, f);
      const legacyNoTv = normalize_(f) === 'no tv here' && Object.prototype.hasOwnProperty.call(data, 'noTvHere');
      if (key === undefined && !legacyNoTv) return; // field not shown for this cabin
      const submitted = normalizeYesNo_(key !== undefined ? data[key] : data.noTvHere);
      const current = normalizeYesNo_(row[col(f)]);
      const bk = findKey(baseChecks, f);
      const base = bk !== undefined ? normalizeYesNo_(baseChecks[bk]) : null;
      if (base !== null) {
        if (submitted !== base) {
          if (current !== base && current !== submitted) conflicts.push(f);
          else if (current !== submitted) changed.push({ field: f, value: submitted });
        }
      } else if (current !== submitted) {
        changed.push({ field: f, value: submitted });
      }
    });

    // Remarks (stored in the Comments column) for every PAX and Crew cabin.
    const cCol = col('Comments');
    if (cCol !== undefined && Object.prototype.hasOwnProperty.call(data, 'comments')) {
      const submitted = String(data.comments == null ? '' : data.comments).slice(0, 2000);
      const current = String(row[cCol] == null ? '' : row[cCol]);
      const base = Object.prototype.hasOwnProperty.call(data, 'baseComments') ? String(data.baseComments == null ? '' : data.baseComments) : current;
      if (submitted !== base) {
        if (current !== base && current !== submitted) conflicts.push('Remarks');
        else if (current !== submitted) changed.push({ field: 'Comments', value: submitted });
      }
    }

    // What to Check — the job list for this cabin, shown under Remarks.
    const wCol = col('What to Check');
    if (wCol !== undefined && Object.prototype.hasOwnProperty.call(data, 'whatToCheck')) {
      const submitted = String(data.whatToCheck == null ? '' : data.whatToCheck).slice(0, 1000);
      const current = String(row[wCol] == null ? '' : row[wCol]);
      const base = Object.prototype.hasOwnProperty.call(data, 'baseWhatToCheck') ? String(data.baseWhatToCheck == null ? '' : data.baseWhatToCheck) : current;
      if (submitted !== base) {
        if (current !== base && current !== submitted) conflicts.push('What to Check');
        else if (current !== submitted) changed.push({ field: 'What to Check', value: submitted });
      }
    }

    // CREW: TV installed, Power Connected and LAN Connected always take the same value (per TV).
    if (!isPax) {
      for (let n = 1; n <= 3; n++) {
        const group = crewLinkedFields_(n);
        const inGroup = f => group.some(g => normalize_(g) === normalize_(f));
        const hit = changed.find(x => inGroup(x.field));
        if (!hit) continue;
        for (let i = changed.length - 1; i >= 0; i--) if (inGroup(changed[i].field)) changed.splice(i, 1);
        group.forEach(f => {
          if (col(f) === undefined) return;
          if (normalizeYesNo_(row[col(f)]) !== hit.value) changed.push({ field: f, value: hit.value });
        });
      }
    }

    // Checklist lock (set by Admin): other technicians cannot untick an item already saved as YES.
    if ((isChecklistLocked_() || normalizeYesNo_(col('Locked') === undefined ? '' : row[col('Locked')]) === 'Yes') && auth.username !== 'Admin') {
      const checklist = {};
      ALL_TV_FIELDS.concat(CABIN_EXTRA_FIELDS).forEach(f => checklist[normalize_(f)] = true);
      const blocked = changed.filter(x => checklist[normalize_(x.field)] && x.value === 'No' && normalizeYesNo_(row[col(x.field)]) === 'Yes').map(x => x.field);
      if (blocked.length) return { success: false, locked: true, cabin, cabinType: type,
        message: 'Admin has locked ticked items. You cannot untick: ' + blocked.join(', ') + '. Nothing was saved — tick it again or ask Admin.' };
    }

    if (conflicts.length) {
      queueTrackerEmail_('CONFLICTING CABIN CHANGES', {
        technician: auth.username, cabin: cabin, cabinType: type,
        details: 'Conflicting field(s): ' + conflicts.join(', ') + '. The save was rejected to avoid overwriting another change.',
        timestamp: formatDate_(new Date())
      });
      return { success: false, conflict: true, cabin, cabinType: type, conflictFields: conflicts,
        message: 'Someone changed ' + conflicts.join(', ') + ' for cabin ' + cabin + '. Refresh this cabin and review the latest values before saving.' };
    }

    const updatedRow = row.slice();
    changed.forEach(x => { updatedRow[col(x.field)] = x.value; });

    const aCol = col('Assigned To');
    if (aCol !== undefined && Object.prototype.hasOwnProperty.call(data, 'assignedTo')) {
      const assigned = String(data.assignedTo || '').trim();
      if (assigned && !AUTHORIZED_TECHNICIANS.includes(assigned)) throw new Error('Choose a valid technician from the list.');
      if (assigned !== assignedOwner && !canManageAssignments_(auth.username)) {
        return { success: false, forbidden: true, message: 'Only Tech 3 or Admin can change cabin technician assignments.' };
      }
      if (assigned !== assignedOwner) updatedRow[aCol] = assigned;
    }

    const now = new Date();
    updatedRow[col('Check By')] = auth.username;
    const tvCount = parseTvCount_(updatedRow[col('TV Count')]) || getTvCount_(type, cabin);
    const status = calculateCabinStatusFromFields_(f => updatedRow[col(f)], tvCount, type);
    updatedRow[col('Status')] = status;
    updatedRow[col('Checked At')] = now;
    sh.getRange(rn, 1, 1, sh.getLastColumn()).setValues([updatedRow]);
    bumpDataVersion_();

    const assignmentChanged = aCol !== undefined && techName_(row[aCol]) !== techName_(updatedRow[aCol]);
    const parts = changed.map(x => (x.field === 'Comments' ? 'Remarks' : x.field) + ': ' + x.value);
    if (assignmentChanged) parts.push('Assigned To: ' + (updatedRow[aCol] || 'Unassigned'));
    const changeText = parts.join('; ') || 'Technician / save timestamp updated';
    appendAudit_(now, auth.username, type, cabin, changeText, status, String(auth.sessionId || ''));
    queueTrackerEmail_('CABIN CHECKLIST UPDATED', { technician: auth.username, cabin, cabinType: type, details: changeText, status, timestamp: formatDate_(now) });

    const merged = !!baseStamp && baseStamp !== currentStamp;
    const response = {
      success: true, rowNumber: rn, cabin, cabinType: type, status, checkBy: auth.username, checkedAt: formatDate_(now), merged,
      assignedTo: aCol === undefined ? '' : techName_(updatedRow[aCol]),
      comments: String(updatedRow[cCol] == null ? '' : updatedRow[cCol]), tvCount,
      whatToCheck: wCol === undefined ? '' : String(updatedRow[wCol] == null ? '' : updatedRow[wCol]),
      noTvHere: normalizeYesNo_(updatedRow[col('No TV Here')]) === 'Yes',
      dataVersion: getDataVersion(),
      message: type + ' cabin ' + cabin + ' saved successfully.' + (merged ? ' Latest non-conflicting changes were merged.' : '')
    };
    if (cacheKey) { try { cache.put(cacheKey, JSON.stringify(response), 21600); } catch (_) {} }
    return response;
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Bulk-assign cabins to one technician (Admin only).
 * `items` may be [{rowNumber, cabin, cabinType}, …] (preferred — safe if rows moved) or plain row numbers.
 */
function assignCabinsToTechnician(token, sessionId, items, technician) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can manage bulk cabin assignments.' };
  const target = String(technician || '').trim();
  if (target && !AUTHORIZED_TECHNICIANS.includes(target)) return { success: false, message: 'Choose a valid technician.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'Select one or more cabins.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('The tracker is busy. Please try again in a few seconds.');
  try {
    const sh = sheet_(), last = sh.getLastRow(), m = headerMap_(sh), assignedIndex = m[normalize_('Assigned To')];
    const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')];
    const all = last >= 2 ? sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues() : [];
    const keyToRow = {};
    all.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c) keyToRow[cabinKey_(r[ti], c)] = i + 2; });
    const rows = [...new Set(items.map(it => (it && typeof it === 'object' && it.cabin)
        ? (keyToRow[cabinKey_(it.cabinType, it.cabin)] || 0)
        : Number(it && typeof it === 'object' ? it.rowNumber : it))
      .filter(n => Number.isInteger(n) && n >= 2 && n <= last))];
    const missing = items.length - rows.length;
    if (!rows.length) return { success: false, message: 'None of the selected cabins are in the sheet any more. Refresh and try again.' };

    const now = new Date(), updates = [], audit = [];
    rows.forEach(rn => {
      const row = all[rn - 2];
      const cabin = String(row[ci] == null ? '' : row[ci]).trim(), type = inferType_(row[ti], cabin);
      if (!cabin) return;
      if (techName_(row[assignedIndex]) === target) return;
      row[assignedIndex] = target;
      updates.push({ rowNumber: rn, cabin, cabinType: type, assignedTo: target });
      audit.push([now, auth.username, type, cabin, 'Assigned To: ' + (target || 'Unassigned'), 'Assignment Updated', String(auth.sessionId || '')]);
    });
    if (updates.length) {
      sh.getRange(2, assignedIndex + 1, all.length, 1).setValues(all.map(r => [r[assignedIndex]]));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
    }
    return { success: true, updated: updates.length, missing, assignments: updates, dataVersion: getDataVersion(),
      message: 'Cabin assignments saved.' + (missing ? ' ' + missing + ' selected cabin(s) were no longer in the sheet.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

/**
 * EXCEL ASSIGNMENT — assigns each cabin to the technician named in the uploaded Excel / CSV list.
 * `items` = [{ cabin, cabinType, technician }, …]   technician '' = unassign. Admin only.
 */
function importUploadedAssignments(token, sessionId, items) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can save cabin assignments from an Excel list.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'No assignments to save.' };
  if (items.length > 5000) return { success: false, message: 'Too many rows in one upload (maximum 5000).' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = sheet_(), m = headerMap_(sh), last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'The tracker sheet is empty.' };
    const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')];
    const aCol = m[normalize_('Assigned To')], wCol = m[normalize_('What to Check')];
    const all = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const keyToRow = {};
    all.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c) keyToRow[cabinKey_(r[ti], c)] = i + 2; });

    const now = new Date(), updates = [], audit = [];
    let missing = 0, invalid = 0, unchanged = 0;
    items.forEach(it => {
      if (!it || typeof it !== 'object') return;
      const hasTech = Object.prototype.hasOwnProperty.call(it, 'technician') && it.technician !== null;
      const hasWhat = Object.prototype.hasOwnProperty.call(it, 'whatToCheck') && it.whatToCheck !== null;
      const raw = hasTech ? String(it.technician).trim() : '';
      const target = raw ? (accountName_(raw) || '') : '';
      if (hasTech && raw && !AUTHORIZED_TECHNICIANS.includes(target)) { invalid++; return; }
      const rn = keyToRow[cabinKey_(it.cabinType, it.cabin)];
      if (!rn) { missing++; return; }
      const row = all[rn - 2];
      const cabin = String(row[ci]).trim(), type = inferType_(row[ti], cabin);
      const parts = [];
      if (hasTech && techName_(row[aCol]) !== target) { row[aCol] = target; parts.push('Assigned To: ' + (target || 'Unassigned')); }
      const what = hasWhat ? String(it.whatToCheck).trim().slice(0, 1000) : '';
      if (hasWhat && wCol !== undefined && String(row[wCol] == null ? '' : row[wCol]).trim() !== what) { row[wCol] = what; parts.push('What to Check: ' + (what || '(cleared)')); }
      if (!parts.length) { unchanged++; return; }
      updates.push({ rowNumber: rn, cabin, cabinType: type, assignedTo: techName_(row[aCol]),
        whatToCheck: wCol === undefined ? '' : String(row[wCol] == null ? '' : row[wCol]) });
      audit.push([now, auth.username, type, cabin, parts.join('; ') + ' (Excel)', 'Assignment Updated', String(auth.sessionId || '')]);
    });

    if (updates.length) {
      sh.getRange(2, aCol + 1, all.length, 1).setValues(all.map(r => [r[aCol]]));
      if (wCol !== undefined) sh.getRange(2, wCol + 1, all.length, 1).setValues(all.map(r => [r[wCol]]));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
      const perTech = {};
      updates.forEach(u => { const t = u.assignedTo || 'Unassigned'; perTech[t] = (perTech[t] || 0) + 1; });
      queueTrackerEmail_('ASSIGNMENTS IMPORTED FROM EXCEL', {
        technician: auth.username,
        details: updates.length + ' cabin(s) updated from an Excel list: ' + Object.keys(perTech).map(t => t + ' ' + perTech[t]).join(', '),
        timestamp: formatDate_(now)
      });
    }
    const notes = [];
    if (unchanged) notes.push(unchanged + ' already up to date');
    if (missing) notes.push(missing + ' not found in the sheet');
    if (invalid) notes.push(invalid + ' with an unknown technician name');
    return { success: true, updated: updates.length, unchanged, missing, invalid, assignments: updates, dataVersion: getDataVersion(),
      message: 'Saved assignments / What to Check for ' + updates.length + ' cabin(s).' + (notes.length ? ' Skipped: ' + notes.join(', ') + '.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

/** What to Check options offered in the assignment Excel drop-down (any other text can be typed). */
const WHAT_TO_CHECK_OPTIONS = ['Full install: TV, power, patch cable/LAN, AP, configure, test',
  'Install TV', 'Connect power', 'Patch cable (PAX) / LAN (Crew)', 'AP check (PAX TV 1)', 'Configure TV',
  'Test TV and channels', 'Recheck held-up issue', 'Verify no TV in cabin', 'Replace damaged TV', 'Final inspection',
  'Security to open the door (PAX)'];

/** Builds the technician assignment Excel from the LIVE sheet (Admin only). Upload it back in the same panel. */
function exportAssignmentTemplate(token) {
  const auth = authenticateToken_(token);
  if (!auth || auth.username !== 'Admin') throw new Error('Only Admin can download the assignment Excel.');
  const records = getData_();
  const fields = ['CABIN TYPE','CABIN NO','DECK','FZ','TV Count','Status','Assigned To','What to Check'];
  const rows = records.map(c => [c.cabinType || '', String(c.cabin || ''), c.deck || '', c.fz || '', c.tvCount || 1,
    c.status || '', c.assignedTo || '', c.whatToCheck || '']);
  let tempId = '';
  try {
    const wb = SpreadsheetApp.create('Technician Assignment');
    tempId = wb.getId();
    const sh = wb.getSheets()[0];
    sh.setName('Assignments');
    const values = [fields].concat(rows);
    sh.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('@');
    sh.getRange(1, 1, values.length, fields.length).setValues(values);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, fields.length).setFontWeight('bold').setFontColor('#ffffff').setBackground('#26174a');
    if (rows.length) {
      sh.getRange(2, 7, rows.length, 2).setBackground('#fff2cc');
      sh.getRange(2, 7, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(AUTHORIZED_TECHNICIANS.concat(['Unassigned']), true).setAllowInvalid(false).build());
      sh.getRange(2, 8, rows.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(WHAT_TO_CHECK_OPTIONS, true).setAllowInvalid(true).build());
      sh.getRange(1, 1, values.length, fields.length).createFilter();
    }
    sh.setColumnWidths(1, 6, 90); sh.setColumnWidth(7, 120); sh.setColumnWidth(8, 380);
    const help = wb.insertSheet('How to use');
    help.getRange(1, 1, 7, 1).setValues([
      ['Technician Cabin Assignment'],
      ['1. Fill in the yellow columns on the Assignments sheet: Assigned To (drop-down) and What to Check (drop-down or type your own).'],
      ['2. Leave Assigned To blank to keep the current technician; choose "Unassigned" to clear it. Clearing What to Check clears it on the cabin.'],
      ['3. Do not change CABIN TYPE or CABIN NO. You may delete rows you are not changing.'],
      ['4. Keep "Assignments" as the first sheet and save as .xlsx.'],
      ['5. Tracker > Technician Cabin Assignments > Upload Assignment Excel, then press "Save Assignments from Excel".'],
      ['Example: PAX | 8100 | 8 | 1 | 1 | Not Started | Tech 1 | Install TV']]);
    help.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    help.setColumnWidth(1, 900);
    SpreadsheetApp.flush();
    const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + encodeURIComponent(tempId) + '/export?format=xlsx', {
      method: 'get', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('Google Sheets XLSX export failed (HTTP ' + resp.getResponseCode() + ').');
    const filename = 'TV_Tracker_Technician_Assignment_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx';
    return { success: true, filename, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: Utilities.base64Encode(resp.getBlob().getBytes()), recordCount: rows.length };
  } finally {
    if (tempId) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {} }
  }
}

/**
 * Adds remarks from an uploaded Excel / CSV cabin list to the Comments (Remarks) column.
 * Existing remarks are kept: the uploaded text is added on a new line, and skipped if the cabin
 * already contains exactly that text. Admin can update any cabin; others only their assigned cabins.
 * `items` = [{ cabin, cabinType, remark }, …]
 */
function importUploadedRemarks(token, sessionId, items) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can save remarks from an uploaded Excel cabin list.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'No uploaded remarks to save.' };
  if (items.length > 5000) return { success: false, message: 'Too many rows in one upload (maximum 5000).' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = sheet_(), m = headerMap_(sh), last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'The tracker sheet is empty.' };
    const col = n => m[normalize_(n)];
    const ci = col('CABIN NO'), ti = col('CABIN TYPE'), cCol = col('Comments'), aCol = col('Assigned To');
    const byCol = col('Check By'), atCol = col('Checked At'), sCol = col('Status');
    const all = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const keyToRow = {};
    all.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c) keyToRow[cabinKey_(r[ti], c)] = i + 2; });

    const now = new Date(), results = [], audit = [];
    let missing = 0, forbidden = 0, unchanged = 0;
    items.forEach(it => {
      const remark = String(it && it.remark != null ? it.remark : '').trim().slice(0, 2000);
      if (!remark) return;
      const rn = keyToRow[cabinKey_(it.cabinType, it.cabin)];
      if (!rn) { missing++; return; }
      const row = all[rn - 2];
      if (auth.username !== 'Admin' && techName_(row[aCol]) !== auth.username) { forbidden++; return; }
      const current = String(row[cCol] == null ? '' : row[cCol]).trim();
      if (current.indexOf(remark) !== -1) { unchanged++; return; }
      const next = (current ? current + '\n' + remark : remark).slice(0, 2000);
      const cabin = String(row[ci]).trim(), type = inferType_(row[ti], cabin);
      row[cCol] = next; row[byCol] = auth.username; row[atCol] = now;
      results.push({ cabin, cabinType: type, comments: next, checkBy: auth.username, checkedAt: formatDate_(now) });
      audit.push([now, auth.username, type, cabin, 'Remarks imported from uploaded list: ' + remark.slice(0, 200), String(row[sCol] || ''), String(auth.sessionId || '')]);
    });

    if (results.length) {
      [cCol, byCol, atCol].forEach(c => sh.getRange(2, c + 1, all.length, 1).setValues(all.map(r => [r[c]])));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
      queueTrackerEmail_('REMARKS IMPORTED FROM UPLOAD', {
        technician: auth.username, details: results.length + ' cabin remark(s) added from an uploaded cabin list.', timestamp: formatDate_(now)
      });
    }
    const notes = [];
    if (unchanged) notes.push(unchanged + ' already had that remark');
    if (missing) notes.push(missing + ' not found in the sheet');
    if (forbidden) notes.push(forbidden + ' not assigned to you');
    return { success: true, updated: results.length, unchanged, missing, forbidden, results, dataVersion: getDataVersion(),
      message: 'Remarks added to ' + results.length + ' cabin(s).' + (notes.length ? ' Skipped: ' + notes.join(', ') + '.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

/**
 * BATCH TICK (Admin only) — ticks (or unticks) the chosen checklist items on many cabins at once,
 * e.g. every cabin from an uploaded Excel / CSV list.
 *   items   = [{ cabin, cabinType }, …]
 *   labels  = checklist items, e.g. ['TV installed','Power Connected']
 *   tvScope = 'all' (TV 1, TV 2, TV 3 of each cabin) or 'tv1' (TV 1 only)
 *   tick    = true to set YES, false to set NO
 * AP and Security to Open the Door are only applied to Guest (PAX) TV 1.
 * CREW: TV installed, Power Connected and LAN Connected are always changed together.
 */
function batchUpdateChecklist(token, sessionId, items, labels, tvScope, tick) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can use Batch Tick Checklist.' };
  const allowed = ['No TV Here','TV installed','Power Connected','Patch cable','AP',SECURITY_DOOR,'Fully Configured','Tested','Held Up'];
  labels = Array.isArray(labels) ? labels.filter(l => allowed.includes(l)) : [];
  if (!labels.length) return { success: false, message: 'Choose at least one checklist item to tick.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'No cabins to update. Upload a cabin list first.' };
  if (items.length > 5000) return { success: false, message: 'Too many cabins in one batch (maximum 5000).' };
  const yesNo = tick ? 'Yes' : 'No', allTvs = String(tvScope || 'all') !== 'tv1';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = sheet_(), m = headerMap_(sh), last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'The tracker sheet is empty.' };
    const col = n => m[normalize_(n)];
    const ci = col('CABIN NO'), ti = col('CABIN TYPE'), aCol = col('Assigned To'), tcCol = col('TV Count');
    const byCol = col('Check By'), atCol = col('Checked At'), sCol = col('Status');
    const all = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const keyToRow = {};
    all.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c) keyToRow[cabinKey_(r[ti], c)] = i + 2; });

    const now = new Date(), results = [], audit = [], seen = {}, touched = {};
    touched[byCol] = touched[atCol] = touched[sCol] = true;
    let missing = 0, forbidden = 0, unchanged = 0;
    items.forEach(it => {
      if (!it || !String(it.cabin || '').trim()) return;
      const key = cabinKey_(it.cabinType, it.cabin);
      if (seen[key]) return;
      seen[key] = true;
      const rn = keyToRow[key];
      if (!rn) { missing++; return; }
      const row = all[rn - 2];
      if (auth.username !== 'Admin' && techName_(row[aCol]) !== auth.username) { forbidden++; return; }
      const cabin = String(row[ci]).trim(), type = inferType_(row[ti], cabin), pax = type === 'PAX';
      const tvCount = parseTvCount_(row[tcCol]) || getTvCount_(type, cabin);
      const n = allTvs ? tvCount : 1, values = {}, changes = [];
      const effLabels = (!pax && labels.some(l => CREW_LINKED_LABELS.includes(l)))
        ? [...new Set(labels.concat(CREW_LINKED_LABELS))] : labels;
      effLabels.forEach(label => {
        for (let k = 1; k <= n; k++) {
          if ((label === 'AP' || label === SECURITY_DOOR) && (!pax || k > 1)) continue; // AP + Security = Guest TV 1 only
          const f = tvFieldName_(k, label), c = col(f);
          if (c === undefined) continue;
          values[f] = !!tick;
          if (normalizeYesNo_(row[c]) !== yesNo) { row[c] = yesNo; touched[c] = true; changes.push(f); }
        }
      });
      if (!changes.length) { unchanged++; return; }
      const status = calculateCabinStatusFromFields_(f => { const c = col(f); return c === undefined ? 'No' : row[c]; }, tvCount, type);
      row[byCol] = auth.username; row[atCol] = now; row[sCol] = status;
      results.push({ cabin, cabinType: type, values, status, checkBy: auth.username, checkedAt: formatDate_(now) });
      audit.push([now, auth.username, type, cabin, 'Batch ' + (tick ? 'ticked' : 'unticked') + ': ' + changes.join(', '), status, String(auth.sessionId || '')]);
    });

    if (results.length) {
      // Only the changed columns are written back, so other columns (and any formulas) are left alone.
      Object.keys(touched).map(Number).forEach(c => sh.getRange(2, c + 1, all.length, 1).setValues(all.map(r => [r[c]])));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
      queueTrackerEmail_('BATCH CHECKLIST UPDATE', {
        technician: auth.username,
        details: (tick ? 'Ticked ' : 'Unticked ') + labels.join(', ') + ' (' + (allTvs ? 'all TVs' : 'TV 1 only') + ') on ' + results.length + ' cabin(s).',
        timestamp: formatDate_(now)
      });
    }
    const notes = [];
    if (unchanged) notes.push(unchanged + ' already set');
    if (missing) notes.push(missing + ' not found in the sheet');
    if (forbidden) notes.push(forbidden + ' not assigned to you');
    return { success: true, updated: results.length, unchanged, missing, forbidden, results, dataVersion: getDataVersion(),
      message: (tick ? 'Ticked ' : 'Unticked ') + labels.join(', ') + ' on ' + results.length + ' cabin(s).' + (notes.length ? ' Skipped: ' + notes.join(', ') + '.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

// ---------------------------------------------------------------------------
// CHECKLIST LOCK (Admin) — when ON, Tech 1, Tech 2, Tech 3 and Tech 4 can tick new items but
// cannot untick an item already saved as YES (cabins and public areas). Admin can change everything.
// ---------------------------------------------------------------------------
const CHECKLIST_LOCK_PROPERTY = 'TRACKER_LOCK_TICKED';
function isChecklistLocked_() { return PropertiesService.getScriptProperties().getProperty(CHECKLIST_LOCK_PROPERTY) === '1'; }

function setChecklistLock(token, sessionId, on) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can lock or unlock ticked checklist items.' };
  const props = PropertiesService.getScriptProperties();
  if (on) props.setProperty(CHECKLIST_LOCK_PROPERTY, '1'); else props.deleteProperty(CHECKLIST_LOCK_PROPERTY);
  appendAudit_(new Date(), auth.username, 'SESSION', '', on ? 'Locked ticked checklist items for other technicians' : 'Unlocked ticked checklist items',
    on ? 'Checklist Locked' : 'Checklist Unlocked', String(auth.sessionId || ''));
  bumpDataVersion_();
  return { success: true, locked: !!on, dataVersion: getDataVersion(),
    message: on ? 'Locked: other technicians can no longer untick items saved as YES.' : 'Unlocked: other technicians can tick and untick items again.' };
}

/** Admin: lock / unlock chosen cabins. items = [{ cabin, cabinType }]. Writes Yes / No in the Locked column. */
function setCabinLocks(token, sessionId, items, on) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can lock or unlock cabins.' };
  if (!Array.isArray(items) || !items.length) return { success: false, message: 'Select one or more cabins.' };
  if (items.length > 5000) return { success: false, message: 'Too many cabins in one go (maximum 5000).' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = sheet_(), m = headerMap_(sh), last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'The tracker sheet is empty.' };
    const ci = m[normalize_('CABIN NO')], ti = m[normalize_('CABIN TYPE')], lc = m[normalize_('Locked')];
    if (lc === undefined) return { success: false, message: 'Locked column is missing. Reload the tracker and try again.' };
    const all = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const keyToRow = {};
    all.forEach((r, i) => { const c = String(r[ci] == null ? '' : r[ci]).trim(); if (c) keyToRow[cabinKey_(r[ti], c)] = i + 2; });
    const want = on ? 'Yes' : 'No', now = new Date(), results = [], audit = [];
    let missing = 0, unchanged = 0;
    items.forEach(it => {
      if (!it) return;
      const rn = keyToRow[cabinKey_(it.cabinType, it.cabin)];
      if (!rn) { missing++; return; }
      const row = all[rn - 2], cabin = String(row[ci]).trim(), type = inferType_(row[ti], cabin);
      if (normalizeYesNo_(row[lc]) === want) { unchanged++; results.push({ cabin, cabinType: type, locked: !!on }); return; }
      row[lc] = want;
      results.push({ cabin, cabinType: type, locked: !!on });
      audit.push([now, auth.username, type, cabin, on ? 'Cabin locked (saved YES items cannot be unticked by other technicians)' : 'Cabin unlocked', on ? 'Locked' : 'Unlocked', String(auth.sessionId || '')]);
    });
    if (audit.length) {
      sh.getRange(2, lc + 1, all.length, 1).setValues(all.map(r => [r[lc] === '' || r[lc] == null ? 'No' : r[lc]]));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
    }
    return { success: true, updated: audit.length, unchanged, missing, results, dataVersion: getDataVersion(),
      message: (on ? 'Locked ' : 'Unlocked ') + audit.length + ' cabin(s).' + (unchanged ? ' ' + unchanged + ' already ' + (on ? 'locked' : 'unlocked') + '.' : '') + (missing ? ' ' + missing + ' not found in the sheet.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}

/** Admin: lock / unlock public-area locations by ID (V001 …). Writes Yes / No in the Locked column. */
function setVenueLocks(token, sessionId, ids, on) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' };
  if (String(sessionId || '').trim() !== String(auth.sessionId || '')) return { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' };
  if (auth.username !== 'Admin') return { success: false, forbidden: true, message: 'Only Admin can lock or unlock locations.' };
  if (!Array.isArray(ids) || !ids.length) return { success: false, message: 'No locations chosen.' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('The tracker is busy saving another update. Please try again in a few seconds.');
  try {
    const sh = ensureVenueSheet_(), m = venueHeaderMap_(sh), last = sh.getLastRow();
    if (last < 2) return { success: false, message: 'The Public Area TVs sheet is empty.' };
    const idc = m[normalize_('ID')], lc = m[normalize_('Locked')], vc = m[normalize_('VENUE')], loc = m[normalize_('LOCATION')];
    if (lc === undefined) return { success: false, message: 'Locked column is missing. Reload the tracker and try again.' };
    const all = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const wanted = {}; ids.forEach(x => { wanted[String(x || '').trim().toUpperCase()] = true; });
    const want = on ? 'Yes' : 'No', now = new Date(), results = [], audit = [];
    all.forEach(r => {
      const id = String(r[idc] || '').trim().toUpperCase();
      if (!id || !wanted[id]) return;
      delete wanted[id];
      results.push({ id, locked: !!on });
      if (normalizeYesNo_(r[lc]) === want) return;
      r[lc] = want;
      const where = String(r[vc] || '') + (r[loc] ? ' — ' + r[loc] : '');
      audit.push([now, auth.username, 'VENUE', where + ' (' + id + ')', on ? 'Location locked (saved YES items cannot be unticked by other technicians)' : 'Location unlocked', on ? 'Locked' : 'Unlocked', String(auth.sessionId || '')]);
    });
    if (audit.length) {
      sh.getRange(2, lc + 1, all.length, 1).setValues(all.map(r => [r[lc] == null ? '' : r[lc]]));
      const history = ensureActivitySheets_().history;
      history.getRange(history.getLastRow() + 1, 1, audit.length, 7).setValues(audit);
      bumpDataVersion_();
    }
    const missing = Object.keys(wanted).length;
    return { success: true, updated: audit.length, missing, results, dataVersion: getDataVersion(),
      message: (on ? 'Locked ' : 'Unlocked ') + audit.length + ' location(s).' + (results.length - audit.length ? ' ' + (results.length - audit.length) + ' already ' + (on ? 'locked' : 'unlocked') + '.' : '') + (missing ? ' ' + missing + ' not found.' : '') };
  } finally { if (lock.hasLock()) lock.releaseLock(); }
}


// ---------------------------------------------------------------------------
// REPORTS, BACKUP & MAINTENANCE (2026-10-21) — Admin
// ---------------------------------------------------------------------------
const PROGRESS_SHEET = 'Daily Progress';
const PROGRESS_HEADERS = ['Date','Guest TVs Installed','Guest TVs Needed','Crew TVs Installed','Crew TVs Needed',
  'Cabin TVs Installed','Cabin TVs Needed','Cabins Fully Configured','Cabins Total','Cabins Held Up',
  'Public Area TVs Installed','Public Area TVs Total','Public Area TVs Configured','Public Area TVs Held Up','Recorded At'];
const BACKUP_FOLDER_NAME = 'Tracker Backups';
const BACKUP_PREFIX = 'Tracker backup ';
const BACKUP_KEEP_DAYS = 14;
const SUMMARY_EMAIL_PROPERTY = 'TRACKER_SUMMARY_EMAIL';
const SUMMARY_HOUR_PROPERTY = 'TRACKER_SUMMARY_HOUR';
const LAST_BACKUP_PROPERTY = 'TRACKER_LAST_BACKUP';
const API_KEY_PROPERTY = 'TRACKER_API_KEY';
const MAINT_TRIGGERS_PROPERTY = 'TRACKER_MAINT_TRIGGERS_V1';

function todayStr_(d) { return Utilities.formatDate(d || new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }

/** { auth } for Admin, otherwise { error } ready to return to the browser. */
function requireAdmin_(token, sessionId) {
  const auth = authenticateToken_(token);
  if (!auth) return { error: { success: false, unauthorized: true, message: 'Your editing session expired. Please sign in again.' } };
  if (sessionId !== undefined && String(sessionId || '').trim() !== String(auth.sessionId || '')) {
    return { error: { success: false, unauthorized: true, message: 'This browser session does not match your sign-in. Please sign in again.' } };
  }
  if (auth.username !== 'Admin') return { error: { success: false, forbidden: true, message: 'Only Admin can use this.' } };
  return { auth };
}

// ---- triggers ----
function installMaintenanceTriggers_() {
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'nightlyMaintenance')) {
    ScriptApp.newTrigger('nightlyMaintenance').timeBased().everyDays(1).atHour(23).create();
  }
  installSummaryTrigger_();
}
function installSummaryTrigger_() {
  const hour = Math.min(23, Math.max(0, Number(PropertiesService.getScriptProperties().getProperty(SUMMARY_HOUR_PROPERTY) || 19)));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailySummaryEmail').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySummaryEmail').timeBased().everyDays(1).atHour(hour).create();
}
/** Installs the nightly + summary triggers once, the first time Admin opens the tracker after this update. */
function ensureMaintenanceTriggers_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(MAINT_TRIGGERS_PROPERTY)) return '';
  try { installMaintenanceTriggers_(); props.setProperty(MAINT_TRIGGERS_PROPERTY, new Date().toISOString()); return ''; }
  catch (err) { console.warn('Maintenance triggers not installed: ' + err); return String(err && err.message ? err.message : err); }
}

/** Runs every night at about 23:00 (script time zone). Each step is independent. */
function nightlyMaintenance() {
  const out = [];
  try { recordDailyProgress_(); out.push('progress recorded'); } catch (e) { out.push('progress failed: ' + e); }
  try { const b = backupSpreadsheet_(); out.push('backup ' + b.name + (b.removed ? ', ' + b.removed + ' old removed' : '')); } catch (e) { out.push('backup failed: ' + e); }
  try { out.push(cleanupExpiredSessions_() + ' expired sessions removed'); } catch (e) { out.push('session clean-up failed: ' + e); }
  try { out.push(cleanupEmailQueue_() + ' email queue rows removed'); } catch (e) { out.push('email queue clean-up failed: ' + e); }
  try { out.push(cleanupActiveUsers_() + ' old active-user rows removed'); } catch (e) { out.push('active users clean-up failed: ' + e); }
  console.log('Nightly maintenance: ' + out.join('; '));
  return out.join('; ');
}

// ---- 2 / 3. clean-up ----
function cleanupExpiredSessions_() {
  const props = PropertiesService.getScriptProperties(), all = props.getProperties(), now = Date.now();
  let removed = 0;
  Object.keys(all).forEach(k => {
    if (k.indexOf('TRACKER_AUTH_') !== 0) return;
    let drop = false;
    try { const a = JSON.parse(all[k]); drop = !a || now - Number(a.lastSeenAt || a.createdAt || 0) > AUTH_CACHE_TTL_SECONDS * 1000; }
    catch (_) { drop = true; }
    if (drop) { props.deleteProperty(k); removed++; }
  });
  return removed;
}
function trimSheetRows_(sh, keepRows, width) {
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, width).clearContent();
  if (keepRows.length) sh.getRange(2, 1, keepRows.length, width).setValues(keepRows);
  const spare = sh.getMaxRows() - (keepRows.length + 1);
  if (spare > 100) sh.deleteRows(keepRows.length + 2 + 100, spare - 100);
}
function cleanupEmailQueue_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMAIL_QUEUE_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues(), now = Date.now();
  const age = v => { const t = v instanceof Date ? v.getTime() : new Date(v).getTime(); return Number.isFinite(t) ? now - t : 0; };
  const keep = rows.filter(r => !((r[8] && age(r[8]) > 3 * 86400000) || (!r[8] && age(r[0]) > 7 * 86400000)) && r.some(String));
  if (keep.length === rows.length) return 0;
  trimSheetRows_(sh, keep, 9);
  return rows.length - keep.length;
}
function cleanupActiveUsers_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_USERS_SHEET);
  if (!sh || sh.getLastRow() < 2) return 0;
  const width = Math.max(6, sh.getLastColumn()), rows = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues(), now = Date.now();
  const keep = rows.filter(r => { const t = r[3] instanceof Date ? r[3].getTime() : new Date(r[3]).getTime(); return r[0] && Number.isFinite(t) && now - t <= 2 * 86400000; });
  if (keep.length === rows.length) return 0;
  trimSheetRows_(sh, keep, width);
  return rows.length - keep.length;
}

// ---- 4. backup ----
function backupSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties();
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
  SpreadsheetApp.flush();
  const name = BACKUP_PREFIX + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
  const cutoff = Date.now() - BACKUP_KEEP_DAYS * 86400000, files = folder.getFiles();
  let removed = 0;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf(BACKUP_PREFIX) === 0 && f.getId() !== copy.getId() && f.getDateCreated().getTime() < cutoff) { f.setTrashed(true); removed++; }
  }
  const info = { name, url: copy.getUrl(), at: formatDate_(new Date()), folderUrl: folder.getUrl() };
  props.setProperty(LAST_BACKUP_PROPERTY, JSON.stringify(info));
  return Object.assign({ removed }, info);
}

// ---- 9. daily progress ----
function progressSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PROGRESS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROGRESS_SHEET);
    sh.getRange(1, 1, 1, PROGRESS_HEADERS.length).setValues([PROGRESS_HEADERS]).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function heldUpCabin_(c) { for (let k = 1; k <= Math.min(3, Math.max(1, Number(c.tvCount) || 1)); k++) if (c[tvFieldName_(k, 'Held Up')]) return true; return false; }
/** Writes (or overwrites) today's row in "Daily Progress". */
function recordDailyProgress_() {
  const cabins = getData_(), s = buildStats_(cabins);
  let v = null; try { v = buildVenueStats_(getVenueData_()); } catch (_) {}
  const day = todayStr_();
  const row = [day, s.pax.tvsInstalled, s.pax.tvExpected, s.crew.tvsInstalled, s.crew.tvExpected,
    s.tvsInstalled, s.tvExpected, s.fullyConfigured, s.total, cabins.filter(heldUpCabin_).length,
    v ? v['Installed'] : '', v ? v.tvs : '', v ? v['Configured'] : '', v ? v['Held Up'] : '', new Date()];
  const sh = progressSheet_(), last = sh.getLastRow();
  sh.getRange('A:A').setNumberFormat('@');
  let rn = 0;
  if (last >= 2) { const days = sh.getRange(2, 1, last - 1, 1).getDisplayValues(); for (let i = days.length - 1; i >= 0; i--) if (days[i][0] === day) { rn = i + 2; break; } }
  if (!rn) rn = last + 1;
  sh.getRange(rn, 1, 1, row.length).setValues([row]);
  return row;
}
function getProgressRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROGRESS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(sh.getLastRow() - 1, 180), start = sh.getLastRow() - n + 1;
  return sh.getRange(start, 1, n, PROGRESS_HEADERS.length).getValues().map(r => ({
    date: String(r[0] instanceof Date ? todayStr_(r[0]) : r[0]),
    paxInstalled: Number(r[1]) || 0, paxNeeded: Number(r[2]) || 0, crewInstalled: Number(r[3]) || 0, crewNeeded: Number(r[4]) || 0,
    cabinInstalled: Number(r[5]) || 0, cabinNeeded: Number(r[6]) || 0, fullyConfigured: Number(r[7]) || 0, cabinsTotal: Number(r[8]) || 0,
    cabinsHeldUp: Number(r[9]) || 0, venueInstalled: Number(r[10]) || 0, venueTotal: Number(r[11]) || 0, venueConfigured: Number(r[12]) || 0, venueHeldUp: Number(r[13]) || 0
  })).filter(x => x.date);
}

// ---- 7. daily summary email ----
function summaryRecipient_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(SUMMARY_EMAIL_PROPERTY) || '';
}
function htmlEsc_(v) { return String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]); }
function buildDailySummary_() {
  const tz = Session.getScriptTimeZone(), today = recordDailyProgress_()[0];
  const cabins = getData_(), s = buildStats_(cabins), venues = (() => { try { return getVenueData_(); } catch (_) { return []; } })();
  const v = buildVenueStats_(venues);
  const prog = getProgressRows_(), prev = prog.filter(x => x.date < today).pop() || null;
  const sinceCab = prev ? s.tvsInstalled - prev.cabinInstalled : null, sinceVen = prev ? v['Installed'] - prev.venueInstalled : null;
  // Per-technician activity today, from the update history.
  const per = {};
  try {
    const h = ensureActivitySheets_().history, last = h.getLastRow();
    if (last >= 2) {
      const n = Math.min(last - 1, 5000);
      h.getRange(last - n + 1, 1, n, 6).getValues().forEach(r => {
        const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
        if (!(d instanceof Date) || isNaN(d) || Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== today) return;
        const tech = String(r[1] || ''), type = String(r[2] || ''), cabin = String(r[3] || ''), ch = String(r[4] || '');
        if (!tech || tech === 'Script' || type === 'SESSION') return;
        const x = per[tech] || (per[tech] = { updates: 0, places: {}, installs: 0, configured: 0 });
        x.updates++; x.places[type + ':' + cabin] = true;
        if (!/^Batch unticked/i.test(ch)) {
          x.installs += (ch.match(/\bTV (?:[23] )?installed(?!: No)/gi) || []).length + (type === 'VENUE' ? (ch.match(/\bInstalled: Yes/g) || []).length : 0);
          x.configured += (ch.match(/\b(?:TV [23] )?Fully Configured(?!: No)/g) || []).length + (type === 'VENUE' ? (ch.match(/\bConfigured: Yes/g) || []).length : 0);
        }
      });
    }
  } catch (_) {}
  const heldCabins = cabins.filter(heldUpCabin_), heldVenues = venues.filter(x => x.checks['Held Up']);
  const decks = Object.keys(s.decks).sort((a, b) => Number(a) - Number(b))
    .map(d => ({ d, left: Math.max(0, (s.decks[d].tvExpected || 0) - (s.decks[d].tvInstalled || 0)), exp: s.decks[d].tvExpected || 0, ins: s.decks[d].tvInstalled || 0 }))
    .filter(x => x.left > 0);
  const pct = (a, b) => b ? Math.round(a / b * 1000) / 10 + '%' : '0%';
  const sign = n => n == null ? 'first day recorded' : (n >= 0 ? '+' : '') + n;
  const td = 'style="padding:6px 10px;border-bottom:1px solid #ddd;text-align:left"', th = 'style="padding:6px 10px;background:#123650;color:#fff;text-align:left"';
  const table = (heads, rows) => '<table style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;margin:6px 0 16px">' +
    '<tr>' + heads.map(x => '<th ' + th + '>' + htmlEsc_(x) + '</th>').join('') + '</tr>' +
    rows.map(r => '<tr>' + r.map(x => '<td ' + td + '>' + htmlEsc_(x) + '</td>').join('') + '</tr>').join('') + '</table>';
  const techRows = Object.keys(per).sort().map(t => [t, Object.keys(per[t].places).length, per[t].updates, per[t].installs, per[t].configured]);
  const heldRows = heldCabins.slice(0, 40).map(c => [c.cabinType + ' ' + c.cabin, 'Deck ' + c.deck, c.assignedTo || 'Unassigned', String(c.comments || '').slice(0, 160)])
    .concat(heldVenues.slice(0, 40).map(x => [x.venue + (x.location ? ' — ' + x.location : ''), 'Deck ' + (x.deck || '?'), x.assignedTo || 'Unassigned', String(x.remarks || '').slice(0, 160)]));
  const html = '<div style="font-family:Arial,sans-serif;color:#111">' +
    '<h2 style="margin:0 0 4px">Daily TV Installation Summary</h2><div style="color:#555;margin-bottom:14px">' + htmlEsc_(today) + ' · times in ' + htmlEsc_(tz) + '</div>' +
    '<h3>Progress</h3>' + table(['', 'Installed', 'Needed', 'Done', 'Since last record'], [
      ['Guest (PAX) cabin TVs', s.pax.tvsInstalled, s.pax.tvExpected, pct(s.pax.tvsInstalled, s.pax.tvExpected), prev ? sign(s.pax.tvsInstalled - prev.paxInstalled) : sign(null)],
      ['Crew cabin TVs', s.crew.tvsInstalled, s.crew.tvExpected, pct(s.crew.tvsInstalled, s.crew.tvExpected), prev ? sign(s.crew.tvsInstalled - prev.crewInstalled) : sign(null)],
      ['All cabin TVs', s.tvsInstalled, s.tvExpected, pct(s.tvsInstalled, s.tvExpected), sign(sinceCab)],
      ['Public area TVs', v['Installed'], v.tvs, pct(v['Installed'], v.tvs), sign(sinceVen)]]) +
    table(['Cabins', 'Count'], [['Fully configured', s.fullyConfigured + ' of ' + s.total], ['In progress', s.inProgress], ['Not started', s.notStarted],
      ['Held Up cabins', heldCabins.length], ['Held Up public area locations', heldVenues.length],
      ['Guest cabins needing Security to open the door', s.pax.securityDoor || 0]]) +
    '<h3>Technicians today</h3>' + (techRows.length ? table(['Technician', 'Cabins / locations updated', 'Saves', 'TV installs ticked', 'Fully configured ticked'], techRows) : '<p>No updates recorded today.</p>') +
    '<h3>TVs still to install per deck</h3>' + (decks.length ? table(['Deck', 'Installed', 'Needed', 'To go'], decks.map(x => ['Deck ' + x.d, x.ins, x.exp, x.left])) : '<p>All cabin TVs are installed.</p>') +
    '<h3>Held Up</h3>' + (heldRows.length ? table(['Cabin / location', 'Deck', 'Assigned to', 'Remarks'], heldRows) +
      (heldCabins.length + heldVenues.length > heldRows.length ? '<p>More in the Held Up report (.xlsx) in the tracker.</p>' : '') : '<p>Nothing is Held Up.</p>') +
    '<p style="color:#777;font-size:12px">Automatic summary from the Cabin TV Tracker. Change the time or address in the tracker: Reports &amp; Maintenance.</p></div>';
  const text = 'Daily TV Installation Summary ' + today + '\n\n' +
    'Cabin TVs installed: ' + s.tvsInstalled + ' of ' + s.tvExpected + ' (' + pct(s.tvsInstalled, s.tvExpected) + '), since last record ' + sign(sinceCab) + '\n' +
    'Public area TVs installed: ' + v['Installed'] + ' of ' + v.tvs + ', since last record ' + sign(sinceVen) + '\n' +
    'Cabins fully configured: ' + s.fullyConfigured + ' of ' + s.total + '\n' +
    'Held Up: ' + heldCabins.length + ' cabins, ' + heldVenues.length + ' public area locations\n' +
    'Guest cabins needing Security to open the door: ' + (s.pax.securityDoor || 0) + '\n\n' +
    'Technicians today:\n' + (techRows.map(r => '- ' + r[0] + ': ' + r[1] + ' cabins/locations, ' + r[3] + ' TV installs ticked').join('\n') || '- no updates') + '\n';
  return { today, html, text };
}
function sendDailySummary_() {
  const to = summaryRecipient_();
  if (!to) return { sent: false, reason: 'No daily summary email address is set.' };
  if (MailApp.getRemainingDailyQuota() <= 0) return { sent: false, reason: 'Google daily email quota is used up; it resets within 24 hours.' };
  const sum = buildDailySummary_();
  MailApp.sendEmail({ to, subject: '[TV Tracker] Daily summary ' + sum.today, htmlBody: sum.html, body: sum.text, name: 'Cabin TV Tracker' });
  return { sent: true, to };
}
/** Trigger handler (daily, at the hour Admin chose). */
function dailySummaryEmail() {
  try { const r = sendDailySummary_(); if (!r.sent) { try { recordDailyProgress_(); } catch (_) {} } console.log('Daily summary: ' + JSON.stringify(r)); }
  catch (err) { console.error('Daily summary failed: ' + err); }
}

// ---- 8. Held Up report ----
function sheetsToXlsx_(title, sheets) {
  let tempId = '';
  try {
    const wb = SpreadsheetApp.create(title);
    tempId = wb.getId();
    sheets.forEach((item, i) => {
      const sh = i === 0 ? wb.getSheets()[0].setName(item.name) : wb.insertSheet(item.name);
      const values = [item.headers].concat(item.rows.length ? item.rows : [item.headers.map((_, k) => k === 0 ? 'None' : '')]);
      sh.getRange(1, 1, values.length, item.headers.length).setValues(values);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, item.headers.length).setFontWeight('bold').setFontColor('#ffffff').setBackground(item.color || '#9f1239');
      if (item.rows.length) sh.getRange(1, 1, values.length, item.headers.length).createFilter();
      sh.autoResizeColumns(1, item.headers.length);
    });
    SpreadsheetApp.flush();
    const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + encodeURIComponent(tempId) + '/export?format=xlsx', {
      method: 'get', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error('Google Sheets XLSX export failed (HTTP ' + resp.getResponseCode() + ').');
    return Utilities.base64Encode(resp.getBlob().getBytes());
  } finally { if (tempId) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {} } }
}
function exportHeldUpReport(token) {
  const a = requireAdmin_(token);
  if (a.error) throw new Error(a.error.message);
  const cabins = getData_().filter(heldUpCabin_);
  const venues = getVenueData_().filter(v => v.checks['Held Up']);
  const cabinRows = cabins.map(c => {
    const n = Math.min(3, Math.max(1, Number(c.tvCount) || 1)), which = [];
    for (let k = 1; k <= n; k++) if (c[tvFieldName_(k, 'Held Up')]) which.push('TV ' + k);
    return [c.cabinType, String(c.cabin), c.deck, c.fz, n, which.join(', '), c.status, c.assignedTo || 'Unassigned', c.comments || '', c.whatToCheck || '', c.checkBy || '', c.checkedAt || ''];
  });
  const venueRows = venues.map(v => [v.venue, v.location, v.deck, v.size ? v.size + '"' : '', v.qty, v.area + ' AREA', v.source, v.status, v.assignedTo || 'Unassigned', v.remarks || '', v.checkBy || '', v.checkedAt || '']);
  const day = todayStr_();
  const base64 = sheetsToXlsx_('Held Up report ' + day, [
    { name: 'Cabins Held Up', headers: ['CABIN TYPE','CABIN NO','DECK','FZ','TV COUNT','HELD UP TVs','STATUS','ASSIGNED TO','REMARKS','WHAT TO CHECK','LAST UPDATED BY','LAST UPDATED AT'], rows: cabinRows },
    { name: 'Public Areas Held Up', headers: ['VENUE','LOCATION','DECK','TV SIZE','QTY','AREA','SOURCE','STATUS','ASSIGNED TO','REMARKS','LAST UPDATED BY','LAST UPDATED AT'], rows: venueRows }]);
  return { success: true, filename: 'TV_Tracker_Held_Up_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', base64, cabins: cabinRows.length, venues: venueRows.length };
}

// ---- web app (Admin) ----
function getAdminStatus(token, sessionId) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  const triggerError = ensureMaintenanceTriggers_();
  const props = PropertiesService.getScriptProperties();
  let backup = null; try { backup = JSON.parse(props.getProperty(LAST_BACKUP_PROPERTY) || 'null'); } catch (_) {}
  const triggers = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  return { success: true, timeZone: Session.getScriptTimeZone(), summaryEmail: props.getProperty(SUMMARY_EMAIL_PROPERTY) || '',
    summaryHour: Number(props.getProperty(SUMMARY_HOUR_PROPERTY) || 19), backup, apiKeySet: !!props.getProperty(API_KEY_PROPERTY),
    nightlyOn: triggers.includes('nightlyMaintenance'), summaryOn: triggers.includes('dailySummaryEmail'), triggerError,
    progress: getProgressRows_(), today: todayStr_() };
}
function saveDailySummarySettings(token, sessionId, email, hour) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  email = String(email || '').trim();
  if (email && !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) return { success: false, message: 'Enter one valid email address (or leave it empty to turn the summary off).' };
  const h = Math.round(Number(hour));
  if (!(h >= 0 && h <= 23)) return { success: false, message: 'Choose a time between 00:00 and 23:00.' };
  const props = PropertiesService.getScriptProperties();
  if (email) props.setProperty(SUMMARY_EMAIL_PROPERTY, email); else props.deleteProperty(SUMMARY_EMAIL_PROPERTY);
  props.setProperty(SUMMARY_HOUR_PROPERTY, String(h));
  try { installSummaryTrigger_(); } catch (err) { return { success: false, message: 'Saved, but the daily timer could not be set: ' + err + '. Use the Tracker Admin menu › Turn on nightly backup… in the spreadsheet.' }; }
  appendAudit_(new Date(), a.auth.username, 'SESSION', '', email ? 'Daily summary set to ' + email + ' at ' + ('0' + h).slice(-2) + ':00' : 'Daily summary turned off', 'Settings', String(a.auth.sessionId || ''));
  return { success: true, message: email ? 'Daily summary will be emailed to ' + email + ' every day at about ' + ('0' + h).slice(-2) + ':00 (' + Session.getScriptTimeZone() + ').' : 'Daily summary turned off.' };
}
function sendDailySummaryNow(token, sessionId) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  const r = sendDailySummary_();
  return r.sent ? { success: true, message: 'Summary sent to ' + r.to + '.', progress: getProgressRows_() } : { success: false, message: r.reason };
}
function backupNow(token, sessionId) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  const b = backupSpreadsheet_();
  appendAudit_(new Date(), a.auth.username, 'SESSION', '', 'Backup made: ' + b.name, 'Backup', String(a.auth.sessionId || ''));
  return { success: true, backup: b, message: 'Backup saved: ' + b.name + (b.removed ? ' (' + b.removed + ' older than ' + BACKUP_KEEP_DAYS + ' days removed)' : '') + '.' };
}
function getOrCreateApiKey(token, sessionId, renew) {
  const a = requireAdmin_(token, sessionId);
  if (a.error) return a.error;
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(API_KEY_PROPERTY);
  if (!key || renew) {
    key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    props.setProperty(API_KEY_PROPERTY, key);
    appendAudit_(new Date(), a.auth.username, 'SESSION', '', renew ? 'Phone app API key renewed' : 'Phone app API key created', 'Settings', String(a.auth.sessionId || ''));
  }
  return { success: true, key, renewed: !!renew };
}

// ---- spreadsheet menu ----
function installMaintenanceFromMenu() {
  installMaintenanceTriggers_();
  PropertiesService.getScriptProperties().setProperty(MAINT_TRIGGERS_PROPERTY, new Date().toISOString());
  SpreadsheetApp.getUi().alert('On: nightly backup, clean-up and daily progress at about 23:00, and the daily summary timer (' + Session.getScriptTimeZone() + ').');
}
function backupNowFromMenu() { const b = backupSpreadsheet_(); SpreadsheetApp.getUi().alert('Backup saved: ' + b.name + '\n' + b.url); }
function setSummaryEmailFromMenu() {
  const ui = SpreadsheetApp.getUi(), props = PropertiesService.getScriptProperties();
  const r = ui.prompt('Daily summary email', 'Email address for the daily summary (empty = off). Current: ' + (props.getProperty(SUMMARY_EMAIL_PROPERTY) || 'none'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const h = ui.prompt('Daily summary time', 'Hour to send it, 0–23 (' + Session.getScriptTimeZone() + '). Current: ' + (props.getProperty(SUMMARY_HOUR_PROPERTY) || '19'), ui.ButtonSet.OK_CANCEL);
  if (h.getSelectedButton() !== ui.Button.OK) return;
  const email = r.getResponseText().trim(), hour = Math.round(Number(h.getResponseText() || 19));
  if (email && !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) { ui.alert('That email address is not valid.'); return; }
  if (!(hour >= 0 && hour <= 23)) { ui.alert('The hour must be 0–23.'); return; }
  if (email) props.setProperty(SUMMARY_EMAIL_PROPERTY, email); else props.deleteProperty(SUMMARY_EMAIL_PROPERTY);
  props.setProperty(SUMMARY_HOUR_PROPERTY, String(hour));
  installSummaryTrigger_();
  ui.alert(email ? 'Daily summary: ' + email + ' at about ' + hour + ':00.' : 'Daily summary turned off.');
}
function sendSummaryNowFromMenu() { const r = sendDailySummary_(); SpreadsheetApp.getUi().alert(r.sent ? 'Summary sent to ' + r.to + '.' : r.reason); }
function showApiKeyFromMenu() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(API_KEY_PROPERTY);
  if (!key) { key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); props.setProperty(API_KEY_PROPERTY, key); }
  SpreadsheetApp.getUi().alert('Phone app API key (put it in API_KEY in the app\'s App.js and keep it private):\n\n' + key);
}


// ---------------------------------------------------------------------------
// CUSTOM EXPORT (Admin) — choose the checklist items, remarks and cabins to export
// opts = { items:[label…], tvScope:'all'|'tv1', scope:'all'|'PAX'|'CREW'|'list', cabins:[{cabin,cabinType}],
//          rows:'all'|'remarks'|'todo'|'done', include:{status,assigned,remarks,what,updated,locked},
//          venues:bool, venueItems:[label…], venueScope:'all'|'list', venueIds:[id…] }
// ---------------------------------------------------------------------------
const CUSTOM_EXPORT_LABELS = { 'No TV Here': 'No TV Here', 'TV installed': 'TV Installed', 'Power Connected': 'Powering On',
  'Patch cable': 'Patch Cable / LAN Connected', 'AP': 'AP (TV 1)', 'Security to Open the Door': 'Security to Open the Door',
  'Tested': 'Tested', 'Fully Configured': 'Fully Configured', 'Held Up': 'Held Up' };

function exportCustomExcel(token, opts) {
  const a = requireAdmin_(token);
  if (a.error) throw new Error(a.error.message);
  opts = opts || {};
  const inc = opts.include || {}, rowsMode = String(opts.rows || 'all');
  const items = (Array.isArray(opts.items) ? opts.items : []).filter(k => CUSTOM_EXPORT_LABELS[k]);
  const allTvs = opts.tvScope !== 'tv1';
  const sheets = [];
  let cabinCount = 0, venueCount = 0;
  const cabinOnly = label => label === 'AP' || label === SECURITY_DOOR; // one per cabin, PAX only

  // ---- cabins
  if (items.length || inc.remarks || inc.what) {
    let cabins = getData_();
    if (opts.scope === 'PAX' || opts.scope === 'CREW') cabins = cabins.filter(c => c.cabinType === opts.scope);
    else if (opts.scope === 'list') {
      const want = {}; (opts.cabins || []).forEach(x => { if (x && x.cabin) want[cabinKey_(x.cabinType, x.cabin)] = true; });
      cabins = cabins.filter(c => want[cabinKey_(c.cabinType, c.cabin)]);
    }
    const cols = [];
    items.forEach(label => {
      const maxN = (cabinOnly(label) || !allTvs) ? 1 : 3;
      for (let n = 1; n <= maxN; n++) {
        const name = CUSTOM_EXPORT_LABELS[label];
        const short = label === 'TV installed' ? 'Installed' : name;
        cols.push({ head: allTvs && !cabinOnly(label) ? 'TV ' + n + ' ' + short : name, field: tvFieldName_(n, label), n, label });
      }
    });
    const headers = ['CABIN TYPE', 'CABIN NO', 'DECK', 'FZ', 'TV COUNT'].concat(cols.map(c => c.head));
    if (inc.status) headers.push('STATUS');
    if (inc.assigned) headers.push('ASSIGNED TO');
    if (inc.locked) headers.push('LOCKED');
    if (inc.what) headers.push('WHAT TO CHECK');
    if (inc.remarks) headers.push('REMARKS');
    if (inc.updated) headers.push('LAST UPDATED BY', 'LAST UPDATED AT');
    const rows = [];
    cabins.forEach(c => {
      const tvCount = Math.min(3, Math.max(1, Number(c.tvCount) || 1)), pax = c.cabinType === 'PAX';
      const applies = col => col.n <= tvCount && !(cabinOnly(col.label) && !pax);
      const vals = cols.map(col => applies(col) ? (c[col.field] ? 'Yes' : 'No') : '');
      // Row filter, using the chosen items only (TV positions marked "No TV Here" are not counted as to-do).
      const work = cols.filter(col => applies(col) && col.label !== 'No TV Here' && col.label !== SECURITY_DOOR && !c[tvFieldName_(col.n, 'No TV Here')]);
      if (rowsMode === 'remarks' && !String(c.comments || '').trim()) return;
      if (rowsMode === 'todo' && !work.some(col => !c[col.field])) return;
      if (rowsMode === 'done' && !(work.length && work.every(col => c[col.field]))) return;
      const r = [c.cabinType, String(c.cabin), c.deck, c.fz, tvCount].concat(vals);
      if (inc.status) r.push(c.status || '');
      if (inc.assigned) r.push(c.assignedTo || 'Unassigned');
      if (inc.locked) r.push(c.locked ? 'Yes' : 'No');
      if (inc.what) r.push(c.whatToCheck || '');
      if (inc.remarks) r.push(c.comments || '');
      if (inc.updated) r.push(c.checkBy || '', c.checkedAt || '');
      rows.push(r);
    });
    cabinCount = rows.length;
    sheets.push({ name: 'Cabins', headers, rows, color: '#123650' });
  }

  // ---- public areas
  if (opts.venues) {
    const vItems = (Array.isArray(opts.venueItems) ? opts.venueItems : []).filter(k => VENUE_CHECKS.includes(k));
    let venues = getVenueData_();
    if (opts.venueScope === 'list') { const want = {}; (opts.venueIds || []).forEach(id => { want[String(id || '').toUpperCase()] = true; }); venues = venues.filter(v => want[v.id]); }
    const headers = ['VENUE', 'LOCATION', 'DECK', 'TV SIZE', 'QTY', 'AREA', 'SOURCE'].concat(vItems.map(k => k.toUpperCase()));
    if (inc.status) headers.push('STATUS');
    if (inc.assigned) headers.push('ASSIGNED TO');
    if (inc.locked) headers.push('LOCKED');
    if (inc.remarks) headers.push('REMARKS');
    if (inc.updated) headers.push('LAST UPDATED BY', 'LAST UPDATED AT');
    const rows = [];
    venues.forEach(v => {
      const work = vItems.filter(k => k !== 'Held Up');
      if (rowsMode === 'remarks' && !String(v.remarks || '').trim()) return;
      if (rowsMode === 'todo' && !work.some(k => !v.checks[k])) return;
      if (rowsMode === 'done' && !(work.length && work.every(k => v.checks[k]))) return;
      const r = [v.venue, v.location, v.deck, v.size ? v.size + '"' : '', v.qty, v.area + ' AREA', v.source].concat(vItems.map(k => v.checks[k] ? 'Yes' : 'No'));
      if (inc.status) r.push(v.status || '');
      if (inc.assigned) r.push(v.assignedTo || 'Unassigned');
      if (inc.locked) r.push(v.locked ? 'Yes' : 'No');
      if (inc.remarks) r.push(v.remarks || '');
      if (inc.updated) r.push(v.checkBy || '', v.checkedAt || '');
      rows.push(r);
    });
    venueCount = rows.length;
    sheets.push({ name: 'Public Areas', headers, rows, color: '#123650' });
  }
  if (!sheets.length) throw new Error('Choose at least one checklist item, Remarks / What to Check, or Public Areas.');
  const base64 = sheetsToXlsx_('custom export ' + todayStr_(), sheets);
  return { success: true, filename: 'TV_Tracker_Custom_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', base64, cabins: cabinCount, venues: venueCount };
}

// ---------------------------------------------------------------------------
// DASHBOARD
// ---------------------------------------------------------------------------
/** A TV is counted ONLY when its "TV Installed" box is ticked (PAX and CREW).
 *  Installed + Fully Configured = done; installed only = in progress; not installed = not started.
 *  TV positions marked "No TV Here" are skipped. */
function countTvs_(c) {
  const n = Math.min(3, Math.max(1, Number(c.tvCount) || 1));
  let expected = 0, installed = 0, noTv = 0, inProgress = 0, done = 0;
  for (let k = 1; k <= n; k++) {
    if (c[tvFieldName_(k, 'No TV Here')]) { noTv++; continue; }
    expected++;
    const isInstalled = !!c[tvFieldName_(k, 'TV installed')];
    const isDone = isInstalled && !!c[tvFieldName_(k, 'Fully Configured')];
    if (isInstalled) installed++;
    if (isDone) done++;
    else if (isInstalled) inProgress++;
  }
  return { slots: n, expected, installed, noTv, inProgress, done, notStarted: expected - installed };
}

function buildStats_(records, subset) {
  const s = { total: records.length, fullyConfigured: 0, inProgress: 0, notStarted: 0, tvInstalled: 0, powerConnected: 0, configured: 0,
    noTvHere: 0, patchCable: 0, apCount: 0, securityDoor: 0, testedCount: 0, heldUpCount: 0, withRemarks: 0,
    tvSlots: 0, tvExpected: 0, tvsInstalled: 0, tvRemaining: 0, tvNoTv: 0,
    tvInProgress: 0, tvDone: 0, tvNotStarted: 0,
    types: {}, decks: {}, technicians: {} };
  records.forEach(c => {
    const tech = String(c.assignedTo || 'Unassigned');
    s.technicians[tech] = (s.technicians[tech] || 0) + 1;
    const t = c.cabinType || 'UNKNOWN', d = String(c.deck || 'Unknown');
    if (!s.types[t]) s.types[t] = { total: 0, fullyConfigured: 0, inProgress: 0, notStarted: 0, tvExpected: 0, tvInstalled: 0 };
    if (!s.decks[d]) s.decks[d] = { total: 0, pax: 0, crew: 0, fullyConfigured: 0, inProgress: 0, notStarted: 0, tvExpected: 0, tvInstalled: 0 };
    s.types[t].total++; s.decks[d].total++;
    if (t === 'PAX') s.decks[d].pax++;
    if (t === 'CREW') s.decks[d].crew++;
    if (c.status === 'Fully Configured') { s.fullyConfigured++; s.types[t].fullyConfigured++; s.decks[d].fullyConfigured++; }
    else if (c.status === 'In Progress') { s.inProgress++; s.types[t].inProgress++; s.decks[d].inProgress++; }
    else { s.notStarted++; s.types[t].notStarted++; s.decks[d].notStarted++; }
    if (c['TV installed']) s.tvInstalled++; // cabins with TV 1 installed (used by the gate cards)
    if (c['Power Connected']) s.powerConnected++;
    if (c['Fully Configured']) s.configured++;
    if (c['No TV Here']) s.noTvHere++;
    if (c['Patch cable']) s.patchCable++; // PAX Patch Cable + CREW LAN Connected
    if (t === 'PAX' && c['AP']) s.apCount++;
    if (t === 'PAX' && c[SECURITY_DOOR]) s.securityDoor++;
    if (c['Tested']) s.testedCount++;
    if (c['Held Up']) s.heldUpCount++;
    if (String(c.comments || '').trim()) s.withRemarks++;
    // Individual TV count (TV 1 + TV 2 + TV 3)
    const tv = countTvs_(c);
    s.tvSlots += tv.slots; s.tvExpected += tv.expected; s.tvsInstalled += tv.installed; s.tvNoTv += tv.noTv;
    s.tvInProgress += tv.inProgress; s.tvDone += tv.done; s.tvNotStarted += tv.notStarted;
    s.types[t].tvExpected += tv.expected; s.types[t].tvInstalled += tv.installed;
    s.decks[d].tvExpected += tv.expected; s.decks[d].tvInstalled += tv.installed;
    s.decks[d].tvInProgress = (s.decks[d].tvInProgress || 0) + tv.inProgress;
  });
  s.tvRemaining = Math.max(0, s.tvExpected - s.tvsInstalled);
  // Furthest gate reached per cabin (for the colour-coded stage bar).
  s.stages = {};
  records.forEach(c => {
    const gates = c.cabinType === 'PAX'
      ? ['TV installed','Power Connected','Patch cable','AP','Tested','Fully Configured']
      : ['TV installed','Power Connected','Patch cable','Tested','Fully Configured'];
    let st = 'none';
    gates.forEach(g => { if (c[g]) st = g; });
    s.stages[st] = (s.stages[st] || 0) + 1;
  });
  s.complete = s.fullyConfigured;
  s.remaining = s.notStarted;
  if (!subset) {
    s.pax = buildStats_(records.filter(c => c.cabinType === 'PAX'), true);
    s.crew = buildStats_(records.filter(c => c.cabinType === 'CREW'), true);
  }
  return s;
}

function getStats() { return buildStats_(getData_()); }

function getTrackerSnapshot(token) {
  const dataVersion = getDataVersion(); // read before the data, so a change in between triggers another reload
  const all = getData_();
  const auth = authenticateToken_(token);
  const cabins = !auth ? [] : (auth.username === 'Admin' ? all : all.filter(c => c.assignedTo === auth.username));
  let venues = [], venueStats = null, venueError = '';
  try { const v = getVenueData_(); venueStats = buildVenueStats_(v); venues = auth ? v : []; }
  catch (err) { venueError = String(err && err.message ? err.message : err); console.warn('Public Area TVs not loaded: ' + venueError); }
  return { cabins, stats: buildStats_(all), venues, venueStats, venueError, lockTicked: isChecklistLocked_(), scriptVersion: SCRIPT_VERSION, serverTime: formatDate_(new Date()), signedInAs: auth ? auth.username : '', dataVersion };
}

// ---------------------------------------------------------------------------
// ACTIVITY, PRESENCE, HISTORY
// ---------------------------------------------------------------------------
function ensureActivitySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let a = ss.getSheetByName(ACTIVE_USERS_SHEET), newA = false;
  if (!a) { a = ss.insertSheet(ACTIVE_USERS_SHEET); newA = true; }
  if (a.getLastRow() === 0) { a.appendRow(['Session ID','Technician','Current Cabin','Last Seen','Connection Status','Work Status']); newA = true; }
  else if (a.getLastColumn() < 6) a.getRange(1, 6).setValue('Work Status');
  let h = ss.getSheetByName(AUDIT_SHEET), newH = false;
  if (!h) { h = ss.insertSheet(AUDIT_SHEET); newH = true; }
  if (h.getLastRow() === 0) { h.appendRow(['Timestamp','Technician','Cabin Type','Cabin No','Changes / Action','Resulting Status','Session ID']); newH = true; }
  if (newA) { a.setFrozenRows(1); a.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff'); }
  if (newH) { h.setFrozenRows(1); h.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#26174a').setFontColor('#ffffff'); }
  return { active: a, history: h };
}

function getActiveUsers(token) {
  if (!authenticateToken_(token)) return [];
  const sh = ensureActivitySheets_().active, last = sh.getLastRow();
  if (last < 2) return [];
  const now = Date.now(), cutoff = 90000;
  const rows = sh.getRange(2, 1, last - 1, 6).getValues();
  const result = [];
  let statusChanged = false;
  rows.forEach(r => {
    const seen = r[3] instanceof Date ? r[3].getTime() : new Date(r[3]).getTime();
    const active = Number.isFinite(seen) && (now - seen) <= cutoff;
    const connection = active ? 'Active' : 'Inactive';
    if (String(r[0] || '') && String(r[1] || '')) {
      result.push({ sessionId: String(r[0]), technician: String(r[1]), cabin: String(r[2] || ''), lastSeen: formatDate_(r[3]),
        active, connectionStatus: connection, workStatus: String(r[5] || 'Working') });
    }
    if (String(r[4] || '') !== connection) { r[4] = connection; statusChanged = true; }
  });
  if (statusChanged) sh.getRange(2, 5, rows.length, 1).setValues(rows.map(r => [r[4]]));
  return result.sort((a, b) => Number(b.active) - Number(a.active) || a.technician.localeCompare(b.technician));
}

function setMyWorkStatus(token, status) {
  const auth = authenticateToken_(token);
  if (!auth) return { success: false, unauthorized: true, message: 'Your session expired. Please sign in again.' };
  status = String(status || '').trim();
  if (!['Working','On Break'].includes(status)) return { success: false, message: 'Choose Working or On Break.' };
  const sh = ensureActivitySheets_().active, last = sh.getLastRow();
  if (last < 2) return { success: false, message: 'No active session found.' };
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === auth.sessionId) {
      sh.getRange(i + 2, 6).setValue(status);
      appendAudit_(new Date(), auth.username, 'SESSION', '', 'Changed work status to ' + status, status, auth.sessionId);
      return { success: true, status };
    }
  }
  return { success: false, message: 'Your session was not found. Please sign in again.' };
}

function kickUser(adminToken, targetSessionId) {
  const admin = authenticateToken_(adminToken);
  if (!admin) return { success: false, unauthorized: true, message: 'Your administrator session expired. Please sign in again.' };
  if (admin.username !== 'Admin') return { success: false, message: 'Permission denied. Only Admin can remove a user session.' };
  targetSessionId = String(targetSessionId || '').trim();
  if (!targetSessionId) return { success: false, message: 'Missing target session.' };
  if (targetSessionId === admin.sessionId) return { success: false, message: 'You cannot kick out your own administrator session.' };
  const props = PropertiesService.getScriptProperties(), all = props.getProperties();
  let target = null, targetKey = '';
  Object.keys(all).forEach(k => {
    if (k.indexOf('TRACKER_AUTH_') === 0) {
      try { const a = JSON.parse(all[k]); if (a && a.sessionId === targetSessionId) { target = a; targetKey = k; } } catch (_) {}
    }
  });
  if (!targetKey) return { success: false, message: 'That user session is already signed out or has expired.' };
  props.deleteProperty(targetKey);
  markSessionInactive_(targetSessionId);
  appendAudit_(new Date(), admin.username, 'SESSION', '', 'Removed ' + target.username + ' from the tracker', 'User Kicked Out', targetSessionId);
  return { success: true, message: target.username + ' has been signed out by Admin.' };
}

function markSessionInactive_(sessionId) {
  const sh = ensureActivitySheets_().active, last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === sessionId) { sh.getRange(i + 2, 3, 1, 3).setValues([['', new Date(Date.now() - 120000), 'Inactive']]); break; }
  }
}

function endUserSession(sessionId, token) {
  sessionId = String(sessionId || '').trim();
  const auth = authenticateToken_(token);
  if (!auth || auth.sessionId !== sessionId) return { success: false, message: 'Session not authorized.' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { success: false };
  try { markSessionInactive_(sessionId); return { success: true }; }
  finally { lock.releaseLock(); }
}

function getUpdateHistory(limit) {
  const sh = ensureActivitySheets_().history, last = sh.getLastRow();
  if (last < 2) return [];
  const n = Math.max(1, Math.min(Number(limit) || 30, 100)), start = Math.max(2, last - n + 1);
  return sh.getRange(start, 1, last - start + 1, 7).getValues().reverse().map(r => ({
    timestamp: formatDate_(r[0]), technician: String(r[1] || ''), cabinType: String(r[2] || ''),
    cabin: String(r[3] || ''), changes: String(r[4] || ''), status: String(r[5] || '')
  }));
}

function appendAudit_(timestamp, technician, type, cabin, changes, status, sessionId) {
  ensureActivitySheets_().history.appendRow([timestamp, technician, type, cabin, changes, status, sessionId]);
}

function testConnection() {
  const sh = sheet_();
  return { sheetName: sh.getName(), cabinRows: Math.max(0, sh.getLastRow() - 1), dataVersion: getDataVersion(),
    message: 'Google Sheets connection is working. The sheet is the live cabin list.' };
}

// ---------------------------------------------------------------------------
// EXCEL EXPORT (Admin only) — includes Remarks
// The web app only shows the Export button to Admin; the check below also blocks anyone else.
// ---------------------------------------------------------------------------
function exportTrackerExcel(token) {
  const auth = authenticateToken_(token);
  if (!auth || auth.username !== 'Admin') throw new Error('Export permission denied. Sign in as Admin to export the complete tracker.');

  const records = getData_();
  const fields = ['CABIN TYPE','CABIN NO','DECK','FZ','TV installed','Power Connected','No TV Here','Patch cable','AP',SECURITY_DOOR,
    'Tested','Fully Configured','Held Up','Check By','Assigned To','Status','Checked At','TV Count']
    .concat(TV2_FIELDS, TV3_FIELDS, PAX_EXTRA_TV_FIELDS, ['What to Check', 'Remarks', 'Locked']);

  const rows = records.map(c => fields.map(label => {
    switch (label) {
      case 'CABIN TYPE': return c.cabinType || '';
      case 'CABIN NO': return c.cabin || '';
      case 'DECK': return c.deck || '';
      case 'FZ': return c.fz || '';
      case 'Check By': return c.checkBy || '';
      case 'Assigned To': return c.assignedTo || '';
      case 'Status': return c.status || '';
      case 'Checked At': return c.checkedAt || '';
      case 'TV Count': return c.tvCount || 1;
      case 'Remarks': return c.comments || '';
      case 'What to Check': return c.whatToCheck || '';
      case 'Locked': return c.locked ? 'Yes' : 'No';
    }
    if (label === SECURITY_DOOR && c.cabinType !== 'PAX') return '';
    return c[label] === true ? 'Yes' : c[label] === false ? 'No' : (c[label] == null ? '' : c[label]);
  }));

  const idx = name => fields.indexOf(name);
  const sheets = [
    { name: 'All Items', data: rows },
    { name: 'In Progress', data: rows.filter(r => r[idx('Status')] === 'In Progress') },
    { name: 'TV Installed', data: rows.filter(r => r[idx('TV installed')] === 'Yes') },
    { name: 'Powering On', data: rows.filter(r => r[idx('Power Connected')] === 'Yes') },
    { name: 'No TV Here', data: rows.filter(r => r[idx('No TV Here')] === 'Yes') },
    { name: 'Tested', data: rows.filter(r => r[idx('Tested')] === 'Yes') },
    { name: 'Fully Configured', data: rows.filter(r => r[idx('Fully Configured')] === 'Yes') },
    { name: 'Held Up', data: rows.filter(r => r[idx('Held Up')] === 'Yes') },
    { name: 'Security to Open Door', data: rows.filter(r => r[idx(SECURITY_DOOR)] === 'Yes') },
    { name: 'With Remarks', data: rows.filter(r => String(r[idx('Remarks')] || '').trim()) }
  ];

  let tempId = '';
  try {
    const workbook = SpreadsheetApp.create('TV Tracker Export');
    tempId = workbook.getId();
    const first = workbook.getSheets()[0];
    first.setName(sheets[0].name);
    sheets.forEach((item, i) => {
      const sh = i === 0 ? first : workbook.insertSheet(item.name);
      const values = [fields].concat(item.data);
      sh.getRange(1, 1, values.length, fields.length).setValues(values);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, fields.length).setFontWeight('bold').setFontColor('#ffffff').setBackground('#123650');
      if (item.data.length) sh.getRange(1, 1, values.length, fields.length).createFilter();
      sh.autoResizeColumns(1, fields.length);
    });
    // Public Area TVs (Crew & Pax areas)
    try {
      const vf = ['VENUE','LOCATION','DECK','TV SIZE','QTY','AREA','SOURCE'].concat(VENUE_CHECKS, ['Status','Assigned To','Check By','Checked At','Remarks']);
      const vrows = getVenueData_().map(v => [v.venue, v.location, v.deck, v.size ? v.size + '"' : '', v.qty, v.area + ' AREA', v.source]
        .concat(VENUE_CHECKS.map(k => v.checks[k] ? 'Yes' : 'No'), [v.status, v.assignedTo, v.checkBy, v.checkedAt, v.remarks]));
      const vsh = workbook.insertSheet('Public Area TVs');
      vsh.getRange(1, 1, vrows.length + 1, vf.length).setValues([vf].concat(vrows));
      vsh.setFrozenRows(1);
      vsh.getRange(1, 1, 1, vf.length).setFontWeight('bold').setFontColor('#ffffff').setBackground('#123650');
      if (vrows.length) vsh.getRange(1, 1, vrows.length + 1, vf.length).createFilter();
      vsh.autoResizeColumns(1, vf.length);
      sheets.push({ name: 'Public Area TVs', data: vrows });
    } catch (err) { console.warn('Public Area TVs export skipped: ' + err); }
    SpreadsheetApp.flush();
    const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + encodeURIComponent(tempId) + '/export?format=xlsx', {
      method: 'get', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('Google Sheets XLSX export failed (HTTP ' + resp.getResponseCode() + ').');
    const filename = 'TV_Tracker_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.xlsx';
    const blob = resp.getBlob().setName(filename);
    return {
      success: true, filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: Utilities.base64Encode(blob.getBytes()),
      recordCount: records.length,
      sheets: sheets.map(s => ({ name: s.name, count: s.data.length }))
    };
  } finally {
    if (tempId) { try { DriveApp.getFileById(tempId).setTrashed(true); } catch (_) {} }
  }
}

// ---------------------------------------------------------------------------
// MOBILE APP JSON API (read-only)
// Called by the Cabin TV Tracker phone app as  <exec URL>?api=sheets&key=...
// The key is created by Admin (Reports & Maintenance › Show API key, or the admin menu) and kept in
// Script Properties. Put the same key in API_KEY in the app's App.js. With no key created, the link is off.
// ---------------------------------------------------------------------------

function handleApi_(e) {
  try {
    const p = e.parameter || {};
    const expectedKey = PropertiesService.getScriptProperties().getProperty(API_KEY_PROPERTY);
    if (!expectedKey) return apiJson_({ ok: false, error: 'Phone app link is turned off. Admin must create the API key in the tracker first.' });
    if (!constantTimeEqual_(String(p.key || ''), expectedKey)) return apiJson_({ ok: false, error: 'Unauthorized' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (p.api === 'sheets') {
      const sheets = ss.getSheets()
        .filter(function (s) { return !s.isSheetHidden(); })
        .map(function (s) {
          return { name: s.getName(), rows: Math.max(s.getLastRow() - 1, 0) };
        });
      return apiJson_({ ok: true, sheets: sheets, updated: new Date().toISOString() });
    }

    if (p.api === 'rows') {
      const sh = ss.getSheetByName(p.sheet);
      if (!sh) return apiJson_({ ok: false, error: 'Sheet not found: ' + p.sheet });
      const values = sh.getDataRange().getDisplayValues();
      if (!values.length) return apiJson_({ ok: true, headers: [], rows: [] });

      var h = 0;
      for (var i = 0; i < Math.min(values.length, 10); i++) {
        if (values[i].filter(String).length >= 2) { h = i; break; }
      }
      const headers = values[h].map(function (v, idx) {
        return String(v).trim() || ('Column ' + (idx + 1));
      });
      const rows = [];
      for (var r = h + 1; r < values.length; r++) {
        const line = values[r];
        if (!line.some(String)) continue;
        const obj = { _row: r + 1 };
        headers.forEach(function (k, c) { obj[k] = line[c]; });
        rows.push(obj);
      }
      return apiJson_({ ok: true, sheet: p.sheet, headers: headers, rows: rows,
                        updated: new Date().toISOString() });
    }
    return apiJson_({ ok: false, error: 'Unknown api: ' + p.api });
  } catch (err) {
    return apiJson_({ ok: false, error: String(err) });
  }
}

function apiJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
