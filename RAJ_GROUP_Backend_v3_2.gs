/**
 * RAJ GROUP Rate Limit Backend v3.2
 * Google Spreadsheet name (recommended): RAJ GROUP Rate Limit Backend
 * Required sheets: CURRENT, HISTORY
 *
 * Deploy as Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 */

const CURRENT_SHEET = 'CURRENT';
const HISTORY_SHEET = 'HISTORY';

const CURRENT_HEADERS = [
  'ID','GrpName','Code','Name','Unit','Type','MRP','Rate','Last Disc %','Scheme Rs','Scheme %','CD %',
  'Min Rate Limit','Remark','Status','Check','Needs Fix','Done Count','Created At','Updated At','Done At','Server Updated At'
];
const HISTORY_HEADERS = ['History Time','History Action'].concat(CURRENT_HEADERS);

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = getOrCreateSheet_(ss, CURRENT_SHEET, CURRENT_HEADERS);
  const history = getOrCreateSheet_(ss, HISTORY_SHEET, HISTORY_HEADERS);
  formatSheet_(current, CURRENT_HEADERS.length);
  formatSheet_(history, HISTORY_HEADERS.length);
  return 'Setup complete';
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || 'current').toLowerCase();
    let result;
    if (action === 'history') {
      const limit = Math.min(Math.max(Number(p.limit || 500), 1), 2000);
      result = { ok: true, items: readHistory_(limit), serverTime: new Date().toISOString() };
    } else if (action === 'ping') {
      result = { ok: true, message: 'RAJ GROUP backend online', serverTime: new Date().toISOString() };
    } else {
      result = { ok: true, items: readCurrent_(), serverTime: new Date().toISOString() };
    }
    return output_(result, p.callback);
  } catch (err) {
    return output_({ ok: false, error: String(err && err.message || err) }, e && e.parameter && e.parameter.callback);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    setup();
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(payload.action || 'update');
    const items = Array.isArray(payload.items) ? payload.items : [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const current = ss.getSheetByName(CURRENT_SHEET);
    const history = ss.getSheetByName(HISTORY_SHEET);

    if (action === 'clear') {
      clearCurrent_(current);
      appendSystemHistory_(history, 'CLEAR ALL');
      SpreadsheetApp.flush();
      return text_({ ok: true, action: action, count: 0 });
    }

    if (action === 'delete') {
      deleteItems_(current, items);
      appendHistory_(history, items, 'DELETE');
      SpreadsheetApp.flush();
      return text_({ ok: true, action: action, count: items.length });
    }

    if (items.length) {
      upsertCurrent_(current, items);
      appendHistory_(history, items, action.toUpperCase());
    }
    SpreadsheetApp.flush();
    return text_({ ok: true, action: action, count: items.length, serverTime: new Date().toISOString() });
  } catch (err) {
    return text_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const existing = sh.getLastColumn() ? sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0] : [];
  let needsHeader = sh.getLastRow() === 0 || String(existing[0] || '').trim() === '';
  if (!needsHeader) {
    for (let i = 0; i < headers.length; i++) {
      if (String(existing[i] || '') !== headers[i]) { needsHeader = true; break; }
    }
  }
  if (needsHeader) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function formatSheet_(sh, cols) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#073f7c').setFontColor('#ffffff');
  sh.getRange('G:L').setNumberFormat('0.00');
  sh.getRange('M:M').setNumberFormat('0.00');
  if (sh.getName() === CURRENT_SHEET) {
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDING / CHECK').setBackground('#fce8e6').setFontColor('#b31412').setRanges([sh.getRange('O2:O')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('DONE').setBackground('#e6f4ea').setFontColor('#137333').setRanges([sh.getRange('O2:O')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('RE-DONE').setBackground('#fef7e0').setFontColor('#b06000').setRanges([sh.getRange('O2:O')]).build()
    ]);
  }
}

function statusOf_(r) {
  if (truthy_(r.needsFix)) return 'PENDING / CHECK';
  if (truthy_(r.done) && truthy_(r.redone)) return 'RE-DONE';
  if (truthy_(r.done)) return 'DONE';
  if (r.updatedAt) return 'PENDING';
  return 'DRAFT';
}

function checkOf_(r) {
  if (String(r.verify || '').toLowerCase() === 'bad') return 'X';
  if (truthy_(r.done) || String(r.verify || '').toLowerCase() === 'ok') return 'OK';
  return '';
}

function rowArray_(r) {
  const now = new Date().toISOString();
  return [
    val_(r.id), val_(r.grp), val_(r.code), val_(r.name), val_(r.unit), val_(r.type || 'WHD'),
    numOrBlank_(r.mrp), numOrBlank_(r.rate), numOrBlank_(r.disc), numOrBlank_(r.schemeRs), numOrBlank_(r.schemePct), numOrBlank_(r.cdPct),
    numOrBlank_(r.minRate), val_(r.remark), statusOf_(r), checkOf_(r), truthy_(r.needsFix) ? 'TRUE' : 'FALSE', Number(r.doneCount || 0),
    val_(r.createdAt), val_(r.updatedAt), val_(r.doneAt), now
  ];
}

function historyRowArray_(r, action) {
  return [new Date().toISOString(), action].concat(rowArray_(r));
}

function upsertCurrent_(sh, items) {
  const lastRow = sh.getLastRow();
  const existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, CURRENT_HEADERS.length).getValues() : [];
  const byId = new Map();
  const byCode = new Map();
  existing.forEach((row, i) => {
    const sheetRow = i + 2;
    if (row[0] !== '') byId.set(String(row[0]), sheetRow);
    if (row[2] !== '') byCode.set(String(row[2]).trim().toUpperCase(), sheetRow);
  });

  const updates = [];
  const appends = [];
  items.forEach(item => {
    const arr = rowArray_(item);
    const id = String(item.id || '');
    const code = String(item.code || '').trim().toUpperCase();
    const target = (id && byId.get(id)) || (code && byCode.get(code));
    if (target) updates.push({ row: target, values: arr, item: item });
    else appends.push({ values: arr, item: item });
  });

  updates.forEach(u => {
    sh.getRange(u.row, 1, 1, CURRENT_HEADERS.length).setValues([u.values]);
    highlightMasterChanges_(sh, u.row, u.item);
  });
  if (appends.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, appends.length, CURRENT_HEADERS.length).setValues(appends.map(x => x.values));
    appends.forEach((a, i) => highlightMasterChanges_(sh, start + i, a.item));
  }
}

function sameNumberOrText_(a, b) {
  const sa = String(a == null ? '' : a).replace(/,/g,'').trim();
  const sb = String(b == null ? '' : b).replace(/,/g,'').trim();
  if (sa === '' && sb === '') return true;
  const na = Number(sa), nb = Number(sb);
  if (sa !== '' && sb !== '' && isFinite(na) && isFinite(nb)) return Math.abs(na - nb) < 0.000001;
  return sa === sb;
}

function highlightMasterChanges_(sh, row, item) {
  // G=MRP, H=Rate, I=Last Disc. Yellow means value differs from uploaded Excel master.
  const rng = sh.getRange(row, 7, 1, 3);
  rng.setBackground('#ffffff');
  const changed = [
    !sameNumberOrText_(item.mrp, item.baseMrp),
    !sameNumberOrText_(item.rate, item.baseRate),
    !sameNumberOrText_(item.disc, item.baseDisc)
  ];
  changed.forEach((isChanged, i) => {
    if (isChanged) sh.getRange(row, 7 + i).setBackground('#fff2cc');
  });
}

function appendHistory_(sh, items, action) {
  if (!items.length) return;
  const values = items.map(r => historyRowArray_(r, action));
  sh.getRange(sh.getLastRow() + 1, 1, values.length, HISTORY_HEADERS.length).setValues(values);
}

function appendSystemHistory_(sh, action) {
  const row = [new Date().toISOString(), action].concat(new Array(CURRENT_HEADERS.length).fill(''));
  sh.getRange(sh.getLastRow() + 1, 1, 1, HISTORY_HEADERS.length).setValues([row]);
}

function deleteItems_(sh, items) {
  if (!items.length || sh.getLastRow() < 2) return;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const ids = new Set(items.map(x => String(x.id || '')));
  const codes = new Set(items.map(x => String(x.code || '').trim().toUpperCase()));
  const rowsToDelete = [];
  vals.forEach((r, i) => {
    if ((r[0] && ids.has(String(r[0]))) || (r[2] && codes.has(String(r[2]).trim().toUpperCase()))) rowsToDelete.push(i + 2);
  });
  rowsToDelete.sort((a,b)=>b-a).forEach(row => sh.deleteRow(row));
}

function clearCurrent_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, CURRENT_HEADERS.length).clearContent();
}

function readCurrent_() {
  setup();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CURRENT_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const v = sh.getRange(2, 1, lastRow - 1, CURRENT_HEADERS.length).getDisplayValues();
  return v.filter(r => r[0] || r[2]).map(currentObject_);
}

function currentObject_(r) {
  const status = r[14] || '';
  const check = r[15] || '';
  return {
    id:r[0], grp:r[1], code:r[2], name:r[3], unit:r[4], type:r[5] || 'WHD', mrp:r[6], rate:r[7], disc:r[8],
    schemeRs:r[9], schemePct:r[10], cdPct:r[11], minRate:r[12], remark:r[13],
    done: status === 'DONE' || status === 'RE-DONE', redone: status === 'RE-DONE', verify: check === 'X' ? 'bad' : (check === 'OK' ? 'ok' : ''),
    needsFix: status === 'PENDING / CHECK', doneCount:Number(r[17] || 0), createdAt:r[18], updatedAt:r[19], doneAt:r[20], serverUpdatedAt:r[21], dirty:false
  };
}

function readHistory_(limit) {
  setup();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const count = Math.min(limit, lastRow - 1);
  const start = lastRow - count + 1;
  const v = sh.getRange(start, 1, count, HISTORY_HEADERS.length).getDisplayValues();
  return v.reverse().map(r => ({
    historyTime:r[0], historyAction:r[1], id:r[2], grp:r[3], code:r[4], name:r[5], unit:r[6], type:r[7], mrp:r[8], rate:r[9], disc:r[10],
    schemeRs:r[11], schemePct:r[12], cdPct:r[13], minRate:r[14], remark:r[15], status:r[16], check:r[17], needsFix:r[18], doneCount:r[19],
    createdAt:r[20], updatedAt:r[21], doneAt:r[22], serverUpdatedAt:r[23]
  }));
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  const cb = String(callback || '');
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function text_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function val_(v) { return v == null ? '' : String(v); }
function numOrBlank_(v) {
  if (v === '' || v == null) return '';
  const n = Number(String(v).replace(/,/g,''));
  return isFinite(n) ? n : '';
}
function truthy_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v) === '1'; }
