// Prep Tracker -> Google Sheet sync.
// Paste into the sheet: Extensions -> Apps Script, replace everything, Save,
// then Deploy -> New deployment -> Web app (Execute as: Me, Who has access: Anyone).

const TASKS = ['move', 'dsa', 'spring', 'sd', 'apply', 'revise', 'plan', 'phone'];
const HEAD = ['Date', 'Gym', 'DSA', 'Spring Boot', 'System design', 'Job hunt',
  'Revision', 'Plan tomorrow', 'Phone away', 'Done', 'Updated'];
const COLS = HEAD.length;

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  if (sh.getFrozenRows() !== 1) {
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLS).setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// GET: return every day so other devices can catch up.
function doGet() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const days = {}, updated = {};
  if (n > 0) {
    sh.getRange(2, 1, n, COLS).getValues().forEach(r => {
      const k = dateKey_(r[0]);
      if (!k) return;
      days[k] = TASKS.filter((t, i) => r[i + 1] === true);
      updated[k] = Number(new Date(r[COLS - 1])) || 0;
    });
  }
  return json_({ ok: true, days, updated });
}

// POST: {items:[{date, done:[ids], updated:ms}]} -> upsert one row per date.
function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const items = body.items || [];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    const keys = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => dateKey_(r[0])) : [];
    const stamps = n > 0 ? sh.getRange(2, COLS, n, 1).getValues().map(r => Number(new Date(r[0])) || 0) : [];
    items.forEach(it => {
      const done = new Set(it.done || []);
      const row = [it.date].concat(TASKS.map(t => done.has(t)), [done.size + '/' + TASKS.length, new Date(it.updated)]);
      const idx = keys.indexOf(it.date);
      if (idx === -1) {
        sh.appendRow(row);
        const r = sh.getLastRow();
        sh.getRange(r, 2, 1, TASKS.length).insertCheckboxes();
        sh.getRange(r, 2, 1, TASKS.length).setValues([row.slice(1, 1 + TASKS.length)]);
        keys.push(it.date); stamps.push(it.updated);
      } else if (it.updated >= stamps[idx]) {
        const r = idx + 2;
        sh.getRange(r, 2, 1, TASKS.length).insertCheckboxes();
        sh.getRange(r, 1, 1, COLS).setValues([row]);
        stamps[idx] = it.updated;
      }
    });
    sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), COLS).sort({ column: 1, ascending: false });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, saved: items.length });
}
