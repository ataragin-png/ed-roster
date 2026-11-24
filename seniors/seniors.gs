/**
 * שרת לשיבוץ רופאים בכירים ברפואה דחופה (גרסה 1)
 * פרויקט Apps Script עצמאי שכותב לגיליון Google נפרד (SHEET_ID). לשוניות נוצרות אוטומטית: requests, state, published, config.
 * פריסה: לפריסה → פריסה חדשה → אפליקציית אינטרנט · להפעיל בתור: אני · למי יש גישה: כל אחד.
 */
const SECRET = '__SECRET__'; // אותו מפתח שמוזן באתר הניהול
const SHEET_ID = '1qwWbIq40KqcuWgYoKGziemYUh6PZTS4VVAWEUKUUmRg'; // גיליון "שיבוץ בכירים"
const MAIN_ID = '1KCavu2KRgD4DzspCwB391_PZmdQiXh3-Sj3H0HKnVHo'; // הסידור הכללי של המלר"ד
const MAIN_COLS = [3, 4, 8, 12, 16, 18, 20, 21, 31, 32, 33, 34, 36]; // עמודות הבכירים בלבד (A=1)

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const admin = body.key === SECRET;
    if (body.action === 'submit') {
      const q = body.q || {};
      if (!/^\d{4}-\d{2}$/.test(String(q.m || '')) || !q.n) return out_({ ok: false, error: 'bad' });
      if (!admin) {
        const dl = (config_().deadlines || {})[q.m];
        if (dl && Date.now() > new Date(dl).getTime() + 10 * 60000) return out_({ ok: false, error: 'closed' });
      }
      const json = JSON.stringify(q);
      if (json.length > 45000) return out_({ ok: false, error: 'too big' });
      sheet_('requests').appendRow([new Date(), q.type === 'chg' ? 'החלפה' : 'בקשה', q.m, String(q.n).slice(0, 60), json]);
      return out_({ ok: true });
    }
    if (!admin) return out_({ ok: false, error: 'key' });
    if (body.action === 'saveState') {
      const s = sheet_('state');
      if (s.getLastRow() && body.base !== undefined && body.base !== null) {
        const cur = String(s.getRange(1, 1).getValue());
        if (cur && String(body.base) !== cur) return out_({ ok: false, error: 'stale', ts: cur });
      }
      const json = String(body.state || '');
      const ts = String(Date.now());
      const chunks = [];
      for (let i = 0; i < json.length; i += 40000) chunks.push(['~' + json.slice(i, i + 40000)]);
      s.clearContents();
      s.getRange(1, 1, 1, 3).setValues([[ts, chunks.length, new Date()]]);
      if (chunks.length) s.getRange(2, 1, chunks.length, 1).setValues(chunks);
      return out_({ ok: true, ts: ts });
    }
    if (body.action === 'publish') {
      if (!/^\d{4}-\d{2}$/.test(String(body.m || ''))) return out_({ ok: false, error: 'bad' });
      const p = sheet_('published');
      const rows = p.getLastRow() ? p.getRange(1, 1, p.getLastRow(), 3).getValues() : [];
      let r = rows.findIndex(function (x) { return String(x[0]) === body.m; });
      const cal = String(body.cal || '');
      if (cal.length > 48000) return out_({ ok: false, error: 'too big' });
      if (r < 0) { p.appendRow(["'" + body.m, cal, new Date()]); }
      else p.getRange(r + 1, 1, 1, 3).setValues([["'" + body.m, cal, new Date()]]);
      return out_({ ok: true });
    }
    if (body.action === 'config') {
      const c = sheet_('config');
      c.clearContents();
      c.getRange(1, 1, 1, 2).setValues([[JSON.stringify(body.config || {}), new Date()]]);
      return out_({ ok: true });
    }
    if (body.action === 'writeMain') {
      const sh = SpreadsheetApp.openById(MAIN_ID).getSheetByName(String(body.name || ''));
      if (!sh) return out_({ ok: false, error: 'no tab' });
      const allowed = MAIN_COLS;
      let n = 0;
      (body.cells || []).forEach(function (c) {
        const row = Number(c[0]), col = Number(c[1]);
        if (row >= 1 && row <= sh.getMaxRows() && allowed.indexOf(col) >= 0) { sh.getRange(row, col).setValue(String(c[2] || '')); n++; }
      });
      return out_({ ok: true, n: n });
    }
    if (body.action === 'markDone') {
      const sh = sheet_('requests');
      const row = Number(body.row);
      if (row >= 1 && row <= sh.getLastRow()) sh.getRange(row, 6).setValue(body.done ? 'טופל' : '');
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const a = e.parameter.action;
  if (a === 'config') {
    const c = config_();
    return out_({ ok: true, config: { names: c.names || [], slots: c.slots || [], deadlines: c.deadlines || {} } });
  }
  if (a === 'calendar') {
    const p = sheet_('published');
    const rows = p.getLastRow() ? p.getRange(1, 1, p.getLastRow(), 2).getValues() : [];
    const months = rows.map(function (x) { return String(x[0]); }).sort();
    const m = e.parameter.m || months[months.length - 1];
    const row = rows.find(function (x) { return String(x[0]) === m; });
    return out_({ ok: true, m: m || null, months: months, cal: row ? String(row[1]) : null });
  }
  if (e.parameter.key !== SECRET) return out_({ ok: false, error: 'key' });
  if (a === 'requests') {
    const sh = sheet_('requests');
    const rows = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 6).getValues() : [];
    const items = [];
    rows.forEach(function (r, i) {
      if (!r[4]) return;
      try { items.push({ row: i + 1, at: r[0] ? new Date(r[0]).getTime() : 0, done: r[5] === 'טופל', q: JSON.parse(r[4]) }); } catch (err) {}
    });
    return out_({ ok: true, items: items });
  }
  if (a === 'readMain') {
    const sh = SpreadsheetApp.openById(MAIN_ID).getSheetByName(String(e.parameter.name || ''));
    if (!sh) return out_({ ok: true, values: null });
    return out_({ ok: true, values: sh.getDataRange().getDisplayValues() });
  }
  if (a === 'state') {
    const s = sheet_('state');
    if (!s.getLastRow()) return out_({ ok: true, ts: '', state: null });
    const head = s.getRange(1, 1, 1, 2).getValues()[0];
    const n = Number(head[1]) || 0;
    const parts = n ? s.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).slice(1); }) : [];
    return out_({ ok: true, ts: String(head[0]), state: parts.join('') || null });
  }
  return out_({ ok: false, error: 'action' });
}

function config_() {
  const c = sheet_('config');
  if (!c.getLastRow()) return {};
  try { return JSON.parse(c.getRange(1, 1).getValue()) || {}; } catch (err) { return {}; }
}

function sheet_(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
