/**
 * Flash Me Baby: Google Sheet sync (v2: every word lives in the sheet)
 *
 * Setup
 * 1. Open your sheet > Extensions > Apps Script. Delete any old code and paste this whole file.
 * 2. Change TOKEN below to your own long random secret (20+ letters and numbers).
 * 3. Pick "setup" in the function menu, click Run, and allow access. This creates the tabs.
 * 4. Deploy > New deployment > type "Web app". Execute as: Me. Who has access: Anyone. Deploy.
 * 5. Copy the Web app URL (ends in /exec). Paste URL and token into the app: Settings > Google Sheet sync.
 * After changing this code later: Deploy > Manage deployments > Edit > Version: New version.
 *
 * Tabs
 *  Words    : one row per word, with its progress. Edit freely: change translations, levels,
 *             topics (comma-separated, new names create new topics), or add rows (leave ID empty).
 *             Deleting a row deletes the word in the app.
 *  Stats    : score, streaks and totals (the "state" row is used by the app, do not edit it).
 *  Sessions : a log of every finished round.
 *
 * The URL plus the token give read/write access to these tabs, so keep both private.
 */
const TOKEN = 'CHANGE-ME-to-a-long-random-secret';

const COLS = {
  words:    ['ID', 'English', 'German', 'Article', 'Plural', 'Level', 'Topics', 'Box', 'Seen', 'Correct', 'Wrong', 'Missed', 'Last seen'],
  stats:    ['Key', 'Value'],
  sessions: ['Date', 'Deck', 'Direction', 'Mode', 'Cards', 'Correct', 'Wrong', 'Skipped', 'Points', 'Percent']
};
const TABS = { words: 'Words', stats: 'Stats', sessions: 'Sessions' };

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN.indexOf('CHANGE-ME') === 0) return out_({ error: 'Set your own TOKEN in the script, then deploy a new version' });
  if (p.t !== TOKEN) return out_({ error: 'Wrong token' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    assignIds_();
    const ver = ver_();
    if (ver > 0 && String(p.v) === String(ver)) return out_({ ok: true, same: true, ver: ver });
    const rows = rows_('words').filter(r => r[0] !== '' && r[0] !== null).map(r => r.map(v => v instanceof Date ? v.toISOString() : v));
    const st = rows_('stats').find(r => r[0] === 'state');
    let stats = null;
    try { stats = st && st[1] ? JSON.parse(st[1]) : null; } catch (err) { stats = null; }
    return out_({ ok: true, ver: ver, rows: rows, stats: stats });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (TOKEN.indexOf('CHANGE-ME') === 0) return out_({ error: 'Set your own TOKEN in the script, then deploy a new version' });
    if (b.t !== TOKEN) return out_({ error: 'Wrong token' });
    const before = ver_(), stale = Number(b.ver || 0) !== before;
    if (b.rows && b.rows.length) upsert_(b.rows);
    if (b.del && b.del.length) remove_(b.del);
    if (b.stats) writeStats_(b.stats);
    if (b.sessions && b.sessions.length) {
      const s = tab_('sessions');
      s.getRange(s.getLastRow() + 1, 1, b.sessions.length, COLS.sessions.length).setValues(b.sessions.map(r => fit_(r, COLS.sessions.length)));
    }
    return out_({ ok: true, ver: bump_(), stale: stale });
  } catch (err) {
    return out_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Runs automatically when you edit the sheet by hand, so the app knows to reload. */
function onEdit(e) {
  try { if (e && e.range && e.range.getSheet().getName() === TABS.words) bump_(); } catch (err) {}
}

function assignIds_() {
  const s = tab_('words'), n = s.getLastRow() - 1;
  if (n < 1) return;
  const ids = s.getRange(2, 1, n, 3).getValues();
  let changed = false;
  ids.forEach((r, i) => {
    if (!r[0] && r[1] && r[2]) { s.getRange(i + 2, 1).setValue('u_s' + Date.now().toString(36) + i); changed = true; }
  });
  if (changed) bump_();
}

function upsert_(list) {
  const s = tab_('words'), n = COLS.words.length, last = s.getLastRow() - 1;
  const ids = last > 0 ? s.getRange(2, 1, last, 1).getValues().map(r => String(r[0])) : [];
  const ix = {};
  ids.forEach((id, i) => { if (id) ix[id] = i; });
  const updates = [], adds = [];
  list.forEach(r => {
    r = fit_(r, n);
    const id = String(r[0]);
    if (id in ix) updates.push([ix[id], r]); else { ix[id] = -1; adds.push(r); }
  });
  if (updates.length > 40) {
    const data = s.getRange(2, 1, last, n).getValues();
    updates.forEach(([i, r]) => { data[i] = r; });
    s.getRange(2, 1, last, n).setValues(data);
  } else {
    updates.forEach(([i, r]) => s.getRange(i + 2, 1, 1, n).setValues([r]));
  }
  if (adds.length) s.getRange(s.getLastRow() + 1, 1, adds.length, n).setValues(adds);
}

function remove_(ids) {
  const s = tab_('words'), last = s.getLastRow() - 1;
  if (last < 1) return;
  const del = {};
  ids.forEach(id => { del[String(id)] = true; });
  const col = s.getRange(2, 1, last, 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (del[String(col[i][0])]) s.deleteRow(i + 2);
}

function writeStats_(st) {
  const s = tab_('stats');
  const rows = [
    ['Score', st.score || 0], ['Current streak', st.streak || 0], ['Best streak', st.best || 0],
    ['Total correct', st.ok || 0], ['Total wrong', st.bad || 0], ['Total skipped', st.skip || 0],
    ['Daily streak (days)', st.dstreak || 0], ['Best daily streak', st.dbest || 0], ['Streak freezes', (st.freeze || 0) + '/2'],
    ['Last update', new Date()], ['state', JSON.stringify(st)]
  ];
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).clearContent();
  s.getRange(2, 1, rows.length, 2).setValues(rows);
}

function ver_() { return Number(PropertiesService.getScriptProperties().getProperty('ver') || 0); }
function bump_() { const v = ver_() + 1; PropertiesService.getScriptProperties().setProperty('ver', String(v)); return v; }

function tab_(k) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(TABS[k]);
  if (!s) {
    s = ss.insertSheet(TABS[k]);
    s.getRange(1, 1, 1, COLS[k].length).setValues([COLS[k]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function rows_(k) {
  const s = tab_(k), n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, COLS[k].length).getValues() : [];
}

function fit_(r, n) { r = r.slice(0, n); while (r.length < n) r.push(''); return r; }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor: creates the tabs and asks for permission. */
function setup() {
  Object.keys(TABS).forEach(tab_);
  const old = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  if (old && old.getLastRow() === 0 && SpreadsheetApp.getActiveSpreadsheet().getSheets().length > 1) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(old);
}
