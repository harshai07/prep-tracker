// Prep Tracker -> Google Sheet sync.
// 1. script.new (standalone project) or Sheet: Extensions -> Apps Script. Replace everything with this file. Save.
// 2. Pick "setup" in the function menu and click Run (styles the sheet, asks for permission).
// 3. Deploy -> New deployment -> Web app. Execute as: Me. Who has access: Anyone. Copy the /exec URL.

const TASKS = ['move', 'dsa', 'spring', 'sd', 'apply', 'revise', 'plan', 'phone'];
const LABELS = ['Gym', 'DSA', 'Spring Boot', 'System design', 'Job hunt', 'Revision', 'Plan tomorrow', 'Phone away'];
const HEAD = ['Date'].concat(LABELS, ['Done', 'Updated']);
const COLS = HEAD.length;            // 11: A..K
const INK = '#141a2e', BLUE = '#3355e0', SOFT = '#e4e9fc', LINE = '#dfe2ec';
const FORMAT_VERSION = '2';
const SHEET_ID = '1YPL_Nj0saWbZfgiLAkDH1j0b3dRi7mZPiLU8bkwZ9mI';
const book_ = () => SpreadsheetApp.openById(SHEET_ID);

/* ---------- one-time styling ---------- */
function setup() {
  const ss = book_();
  const log = ss.getSheets().filter(s => s.getName() !== 'Summary')[0];
  log.setName('Log');
  log.getRange(1, 1, 1, COLS).setValues([HEAD]);
  if (log.getMaxColumns() > COLS) log.deleteColumns(COLS + 1, log.getMaxColumns() - COLS);

  log.setFrozenRows(1);
  log.setRowHeight(1, 38);
  log.getRange(1, 1, 1, COLS)
    .setBackground(INK).setFontColor('#ffffff').setFontWeight('bold')
    .setFontFamily('Roboto Mono').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  log.getRange('A:A').setNumberFormat('@').setFontFamily('Roboto Mono');
  log.getRange('B:J').setHorizontalAlignment('center');
  log.getRange('J:J').setNumberFormat('0" / 8"').setFontWeight('bold');
  log.getRange('K:K').setNumberFormat('d mmm, h:mm am/pm').setFontColor('#5f6782');
  log.setColumnWidth(1, 112);
  log.setColumnWidths(2, 8, 104);
  log.setColumnWidth(10, 84);
  log.setColumnWidth(11, 150);
  log.getRange(1, 1, log.getMaxRows(), COLS).setBorder(null, null, true, null, null, true, LINE, SpreadsheetApp.BorderStyle.SOLID);

  const body = log.getRange('A2:K');
  const done = log.getRange('J2:J');
  log.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$J2>=8').setBackground(SOFT).setFontColor(BLUE)
      .setRanges([body]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#fbe3e7', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMidpointWithValue('#fff1d6', SpreadsheetApp.InterpolationType.NUMBER, '4')
      .setGradientMaxpointWithValue('#cdd7fb', SpreadsheetApp.InterpolationType.NUMBER, '8')
      .setRanges([done]).build()
  ]);

  buildSummary_(ss);
  PropertiesService.getScriptProperties().setProperty('fmt', FORMAT_VERSION);
}

function buildSummary_(ss) {
  const sm = ss.getSheetByName('Summary') || ss.insertSheet('Summary', 0);
  sm.clear();
  sm.setHiddenGridlines(true);
  sm.setColumnWidth(1, 24);
  sm.setColumnWidth(2, 170);
  sm.setColumnWidth(3, 110);
  sm.setColumnWidth(4, 90);
  sm.setColumnWidth(5, 260);

  sm.getRange('B2').setValue('Prep Pipeline').setFontSize(22).setFontWeight('bold').setFontColor(INK);
  sm.getRange('B3').setValue('Live from the tracker. Every tick lands in the Log tab.').setFontColor('#5f6782');

  // KPI tiles
  const kpis = [
    ['Days logged', '=COUNTA(Log!A2:A)'],
    ['DSA days', '=COUNTIF(Log!C2:C,TRUE)'],
    ['Full days (6+)', '=COUNTIF(Log!J2:J,">=6")'],
    ['Avg stages / day', '=IFERROR(ROUND(AVERAGE(Log!J2:J),1),0)']
  ];
  kpis.forEach((k, i) => {
    const r = 5 + i;
    sm.getRange(r, 2).setValue(k[0]).setFontColor('#5f6782');
    sm.getRange(r, 3).setFormula(k[1]).setFontSize(16).setFontWeight('bold').setFontColor(BLUE).setHorizontalAlignment('left');
  });

  // last 30 days sparkline
  sm.getRange('B10').setValue('Last 30 days').setFontWeight('bold').setFontColor(INK);
  sm.getRange('B11:E11').merge();
  sm.setRowHeight(11, 70);
  sm.getRange('B11').setFormula(
    '=IFERROR(SPARKLINE(SORT(FILTER(Log!J2:J31,Log!J2:J31<>""),FILTER(Log!A2:A31,Log!J2:J31<>""),TRUE),' +
    '{"charttype","column";"color","' + BLUE + '";"ymin",0;"ymax",8}),"No days yet")');

  // per-stage completion
  sm.getRange('B13:E13').setValues([['Stage', 'Days done', 'Rate', '']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground(INK).setFontFamily('Roboto Mono');
  LABELS.forEach((label, i) => {
    const r = 14 + i;
    const col = String.fromCharCode(66 + i); // B..I on Log
    sm.getRange(r, 2).setValue(label);
    sm.getRange(r, 3).setFormula('=COUNTIF(Log!' + col + '2:' + col + ',TRUE)');
    sm.getRange(r, 4).setFormula('=IFERROR(C' + r + '/COUNTA(Log!A2:A),0)').setNumberFormat('0%');
    sm.getRange(r, 5).setFormula('=SPARKLINE(D' + r + ',{"charttype","bar";"max",1;"color1","' + BLUE + '"})');
    if (i % 2) sm.getRange(r, 2, 1, 4).setBackground('#f4f5f9');
  });
  sm.getRange('C14:D21').setHorizontalAlignment('center');
}

/* ---------- web app ---------- */
function sheet_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('fmt') !== FORMAT_VERSION) setup();
  return book_().getSheetByName('Log');
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}

// GET: every day, so other devices can catch up.
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

// POST: {items:[{date, done:[ids], updated:ms}]} -> upsert one row per date, newest first.
function doPost(e) {
  const items = (JSON.parse((e.postData && e.postData.contents) || '{}').items) || [];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    const keys = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => dateKey_(r[0])) : [];
    const stamps = n > 0 ? sh.getRange(2, COLS, n, 1).getValues().map(r => Number(new Date(r[0])) || 0) : [];
    items.forEach(it => {
      const done = new Set(it.done || []);
      const row = [it.date].concat(TASKS.map(t => done.has(t)), [done.size, new Date(it.updated)]);
      let idx = keys.indexOf(it.date);
      if (idx !== -1 && it.updated < stamps[idx]) return;   // older than what's stored
      if (idx === -1) { keys.push(it.date); stamps.push(0); idx = keys.length - 1; }
      const r = idx + 2;
      sh.getRange(r, 2, 1, TASKS.length).insertCheckboxes();
      sh.getRange(r, 1, 1, COLS).setValues([row]);
      stamps[idx] = it.updated;
    });
    const last = sh.getLastRow();
    if (last > 2) sh.getRange(2, 1, last - 1, COLS).sort({ column: 1, ascending: false });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, saved: items.length });
}
