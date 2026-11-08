/**
 * JSON API for the TV mobile app.
 * Add this as a new file (Api.gs) in your Apps Script project.
 *
 * Then change your existing doGet so it starts like this:
 *
 *   function doGet(e) {
 *     if (e && e.parameter && e.parameter.api) return handleApi_(e);
 *     // ...your existing code that returns Index.html stays below...
 *   }
 */

// Leave blank if the script is bound to the tracker spreadsheet.
// If it's a standalone script, paste the spreadsheet ID here.
const API_SHEET_ID = '';

// Shared secret. Must match API_KEY in the app's App.js.
const API_KEY = 'change-me-to-something-long';

function handleApi_(e) {
  try {
    const p = e.parameter || {};
    if (API_KEY && p.key !== API_KEY) {
      return apiJson_({ ok: false, error: 'Unauthorized' });
    }

    const ss = API_SHEET_ID
      ? SpreadsheetApp.openById(API_SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();

    // List of visible tabs with row counts
    if (p.api === 'sheets') {
      const sheets = ss.getSheets()
        .filter(function (s) { return !s.isSheetHidden(); })
        .map(function (s) {
          return { name: s.getName(), rows: Math.max(s.getLastRow() - 1, 0) };
        });
      return apiJson_({ ok: true, sheets: sheets, updated: new Date().toISOString() });
    }

    // All rows of one tab, as objects keyed by header
    if (p.api === 'rows') {
      const sh = ss.getSheetByName(p.sheet);
      if (!sh) return apiJson_({ ok: false, error: 'Sheet not found: ' + p.sheet });

      const values = sh.getDataRange().getDisplayValues();
      if (!values.length) return apiJson_({ ok: true, headers: [], rows: [] });

      // Header = first row (within the top 10) that has at least 2 filled cells
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
        headers.forEach(function (key, c) { obj[key] = line[c]; });
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
