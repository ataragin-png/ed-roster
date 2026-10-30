/**
 * שרת לשיבוץ מתמחים ברפואה דחופה (גרסה 14)
 * בקשות מהטופס, פרסום הלוח, שמירת כל נתוני האתר, ולשונית נפרדת לכל חודש.
 * גרסה 9: מייל לאחראי על כל בקשת החלפה, עם קישור לאישור באתר השיבוץ.
 * גרסה 10: בדיקה אוטומטית של לוח התורנויות (חפיפות וחוסר מנוחה) עם מייל התראה.
 * גרסה 11: מסגרת אדומה סביב השורה של היום בלוח התורנויות, מתעדכנת כל לילה.
 * גרסה 12: גישה לסידור הכללי של המלר״ד: קריאה, וכתיבה לתאים מסוימים בלבד.
 * גרסה 13: עדכון אוטומטי של עמודות המתמחים בסידור הכללי לפי לוח התורנויות (מנובמבר 2026).
 * גרסה 14: התראה במייל כשמתמחה עובר את מכסת התורנויות או המשמרות שלו בלוח התורנויות.
 */
const SECRET = '__SECRET__'; // אותו מפתח שמוזן באתר השיבוץ
const NOTIFY_EMAIL = 'ataragin@gmail.com'; // לאן נשלחות התראות על בקשות החלפה
const ADMIN_URL = 'https://ataragin-png.github.io/ed-roster/admin.html';
const ALLOWED_IDS = ['1tHLsrjRMdTiqh6kVVxAhoo1DD3oOSdf6LfArxzfiHZI', '1IbVJDhWcut7F37GjC4nFQNwViArzh5nvR1EhfW4eoYQ', '1KCavu2KRgD4DzspCwB391_PZmdQiXh3-Sj3H0HKnVHo']; // תוכנית ההתמחות, לוח התורנויות והסידור הכללי בלבד
const MAIN_ID = '1KCavu2KRgD4DzspCwB391_PZmdQiXh3-Sj3H0HKnVHo'; // הסידור הכללי של המלר״ד

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (body.action === 'submit') {
      const q = body.q || {};
      if (!q.m || !q.n) return out_({ ok: false, error: 'bad' });
      if (q.type !== 'chg') {
        const dl = deadlines_()[q.m];
        if (dl && Date.now() > new Date(dl).getTime() + 10 * 60000) return out_({ ok: false, error: 'closed' });
      }
      const json = JSON.stringify(q);
      if (json.length > 45000) return out_({ ok: false, error: 'too big' });
      const rq = sheet_('requests');
      rq.appendRow([new Date(), q.type === 'chg' ? 'החלפה' : 'בקשה', q.m, String(q.n).slice(0, 60), json]);
      if (q.type === 'chg') {
        const row = rq.getLastRow();
        try { notifyChg_(q, row); } catch (err) { console.error('mail failed: ' + err); }
      }
      return out_({ ok: true });
    }
    if (body.key !== SECRET) return out_({ ok: false, error: 'key' });
    if (body.action === 'publish') {
      const p = sheet_('published');
      p.clear();
      p.getRange(1, 1, 1, 3).setValues([[new Date(), body.m, body.cal]]);
      return out_({ ok: true });
    }
    if (body.action === 'exportBoard') {
      const ss = openAllowed_(body.id);
      const name = String(body.name).slice(0, 60);
      let sh = ss.getSheetByName(name);
      if (sh && !body.overwrite) return out_({ ok: false, error: 'exists' });
      if (!sh) sh = ss.insertSheet(name, ss.getSheets().length);
      sh.clear();
      const rows = body.rows || [];
      const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
      const vals = rows.map(function (r) {
        const out = [];
        for (let i = 0; i < width; i++) {
          let v = r[i] === undefined || r[i] === null ? '' : r[i];
          if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
          out.push(v);
        }
        return out;
      });
      sh.getRange(1, 1, vals.length, width).setValues(vals);
      sh.setRightToLeft(true);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#D9E2F3');
      (body.weekend || []).forEach(function (d) { sh.getRange(d + 1, 1, 1, width).setBackground('#FCE5CD'); });
      sh.getRange(1, 1, vals.length, width).setHorizontalAlignment('center');
      sh.autoResizeColumns(1, width);
      try { if (String(body.id) === BOARD_ID) markToday(); } catch (err) {}
      try { if (String(body.id) === BOARD_ID) syncMain(); } catch (err) {}
      return out_({ ok: true });
    }
    if (body.action === 'writeCells') {
      // כותב רק את התאים שנשלחו: [[שורה, עמודה, ערך], ...]. לא נוגע בשום תא אחר
      const ss = openAllowed_(body.id);
      const sh = ss.getSheetByName(String(body.name));
      if (!sh) return out_({ ok: false, error: 'no sheet' });
      const cells = body.cells || [];
      if (cells.length > 2000) return out_({ ok: false, error: 'too many' });
      cells.forEach(function (c) {
        let v = c[2] === undefined || c[2] === null ? '' : c[2];
        if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
        sh.getRange(Number(c[0]), Number(c[1])).setValue(v);
      });
      return out_({ ok: true, written: cells.length });
    }
    if (body.action === 'saveState') {
      const s = sheet_('state');
      const json = String(body.state || '');
      const chunks = [];
      for (let i = 0; i < json.length; i += 40000) chunks.push(['~' + json.slice(i, i + 40000)]);
      s.clearContents();
      s.getRange(1, 1, 1, 3).setValues([[String(body.ts || Date.now()), chunks.length, new Date()]]);
      if (chunks.length) s.getRange(2, 1, chunks.length, 1).setValues(chunks);
      if (body.tab && body.tab.base && body.tab.rows) writeTab_(body.tab);
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
  if (a === 'calendar') {
    const p = sheet_('published');
    const cal = p.getLastRow() ? p.getRange(1, 3).getValue() : '';
    return out_({ ok: true, cal: cal || null });
  }
  if (a === 'plan') {
    const st = readState_();
    if (!st) return out_({ ok: true, plan: {} });
    const x = JSON.parse(st);
    const names = {};
    (x.people || []).forEach(function (p) { names[p.id] = p.name; });
    const plan = {};
    const rp = x.rotPlan || {};
    Object.keys(rp).forEach(function (id) { if (names[id]) plan[names[id]] = rp[id]; });
    return out_({ ok: true, plan: plan, deadlines: (x.settings && x.settings.req && x.settings.req.deadlines) || {} });
  }
  if (e.parameter.key !== SECRET) return out_({ ok: false, error: 'key' });
  if (a === 'program') {
    const ss = openAllowed_(e.parameter.id);
    const gid = Number(e.parameter.gid || 0);
    const sh = ss.getSheets().filter(function (x) { return x.getSheetId() === gid; })[0] || ss.getSheets()[0];
    const rng = sh.getDataRange();
    const vals = rng.getDisplayValues();
    if (vals.length && vals[0].length) {
      rng.getMergedRanges().forEach(function (m) {
        const r0 = m.getRow() - 1, c0 = m.getColumn() - 1, v = (vals[r0] || [])[c0];
        for (let r = r0; r < r0 + m.getNumRows() && r < vals.length; r++)
          for (let c = c0; c < c0 + m.getNumColumns() && c < vals[r].length; c++) vals[r][c] = v;
      });
    }
    return out_({ ok: true, name: ss.getName(), values: vals });
  }
  if (a === 'listTabs') {
    const ss = openAllowed_(e.parameter.id);
    return out_({ ok: true, name: ss.getName(), tabs: ss.getSheets().map(function (x) { return x.getName(); }) });
  }
  if (a === 'readTab') {
    const ss = openAllowed_(e.parameter.id);
    const sh = ss.getSheetByName(e.parameter.name);
    if (!sh) return out_({ ok: true, values: null });
    const rng = sh.getDataRange();
    const merges = rng.getMergedRanges().map(function (m) { return [m.getRow(), m.getColumn(), m.getNumRows(), m.getNumColumns()]; });
    return out_({ ok: true, values: rng.getDisplayValues(), merges: merges });
  }
  if (a === 'readBoard') {
    const ss = openAllowed_(e.parameter.id);
    const sh = ss.getSheetByName(e.parameter.name);
    if (!sh) return out_({ ok: true, values: null });
    return out_({ ok: true, values: sh.getDataRange().getDisplayValues() });
  }
  if (a === 'requests') {
    const rows = sheet_('requests').getDataRange().getValues().slice(1);
    const items = [];
    rows.forEach(function (r, i) {
      try { items.push({ row: i + 2, q: JSON.parse(r[4]) }); } catch (err) {}
    });
    return out_({ ok: true, items: items });
  }
  if (a === 'state') {
    const s = sheet_('state');
    if (!s.getLastRow()) return out_({ ok: true, state: null });
    return out_({ ok: true, ts: String(s.getRange(1, 1).getValue()), state: readState_() });
  }
  return out_({ ok: true, service: 'ed-roster' });
}

function deadlines_() {
  try {
    const st = readState_();
    if (!st) return {};
    const x = JSON.parse(st);
    return (x.settings && x.settings.req && x.settings.req.deadlines) || {};
  } catch (err) { return {}; }
}

function readState_() {
  const s = sheet_('state');
  if (!s.getLastRow()) return null;
  const n = Number(s.getRange(1, 2).getValue()) || 0;
  if (!n) return null;
  return s.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]).slice(1); }).join('') || null;
}

function writeTab_(tab) {
  const ss = SpreadsheetApp.getActive();
  const base = String(tab.base).slice(0, 40);
  let sh = ss.getSheets().filter(function (x) { return x.getName() === base || x.getName().indexOf(base + ' (') === 0; })[0];
  if (!sh) sh = ss.insertSheet(base);
  const name = String(tab.name).slice(0, 60);
  if (sh.getName() !== name) sh.setName(name);
  sh.clear();
  const rows = tab.rows;
  const width = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
  const vals = rows.map(function (r) {
    const out = [];
    for (let i = 0; i < width; i++) {
      let v = r[i] === undefined || r[i] === null ? '' : r[i];
      if (typeof v === 'string' && /^[=+\-@]/.test(v)) v = "'" + v;
      out.push(v);
    }
    return out;
  });
  sh.getRange(1, 1, vals.length, width).setValues(vals);
  sh.setRightToLeft(true);
  sh.setFrozenRows(2);
  sh.setFrozenColumns(1);
  sh.getRange(1, 1, 2, width).setFontWeight('bold').setBackground('#E3EAF1');
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    if (name === 'requests') s.appendRow(['זמן', 'סוג', 'חודש', 'שם', 'נתונים']);
  }
  return s;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function openAllowed_(id) {
  if (ALLOWED_IDS.indexOf(String(id)) < 0) throw new Error('not allowed');
  return SpreadsheetApp.openById(id);
}


/* ---------- התראה במייל על בקשת החלפה ---------- */
const TNAMES_ = { m: 'בוקר אגף א', n: 'לילה אגף א', b: 'בוקר צד ב', v: 'תורנות צד ב', bv: 'בוקר + תורנות צד ב (סופ״ש)', g: 'בוקר צד ג', p: 'תורנות צד ג (פנימית)', c: 'מהלכים פנימית', mc: 'בוקר מהלכים כירורגיה', nc: 'ערב מהלכים כירורגיה', o: 'אורטופדיה', k: 'ילדים', i: 'טיפול נמרץ', a: 'מד״א', z: 'בוקר אורתופדיה', t: 'סטאז׳ר' };
const WD_ = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש'];

function publishedCal_() {
  try {
    const p = sheet_('published');
    if (!p.getLastRow()) return null;
    const b64 = String(p.getRange(1, 3).getValue() || '');
    if (!b64) return null;
    return JSON.parse(Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8'));
  } catch (err) { return null; }
}

function shiftLabel_(cal, name, d, c, m) {
  let label = TNAMES_[c] || c;
  if (cal && cal.m === m && (c === 'mc' || c === 'nc')) {
    const norm = function (x) { return String(x || '').replace(/[‎‏‪-‮]/g, '').trim(); };
    const key = Object.keys(cal.a || {}).filter(function (k) { return norm(k) === norm(name); })[0];
    const list = key ? cal.a[key] : [];
    const has = function (code) { return list.some(function (x) { return x[0] === d && x[1] === code; }); };
    if (has('mc') && has('nc')) label = '24 שעות כירורגיה (בוקר + ערב, עוברים יחד)';
  }
  const ym = String(m).split('-').map(Number);
  const wd = new Date(ym[0], ym[1] - 1, d).getDay();
  return label + ', ' + d + '.' + ym[1] + ' (יום ' + WD_[wd] + ')';
}

function esc_(x) { return String(x === undefined || x === null ? '' : x).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

function notifyChg_(q, row) {
  if (!NOTIFY_EMAIL) return;
  const cal = publishedCal_();
  const swap = q.kind === 'swap';
  const link = ADMIN_URL + '?chg=' + row;
  const lines = [
    '<b>' + esc_(q.n) + '</b> נותן/ת: ' + esc_(shiftLabel_(cal, q.n, q.d1, q.c1, q.m)),
    swap ? '<b>' + esc_(q.w) + '</b> נותן/ת בתמורה: ' + esc_(shiftLabel_(cal, q.w, q.d2, q.c2, q.m)) : 'המשמרת עוברת ל<b>' + esc_(q.w) + '</b>',
    q.agree ? '<span style="color:#1a7f37">המתמחה השני סימן שהוא מסכים</span>' : '<span style="color:#b35900">לא סומן שהמתמחה השני מסכים</span>'
  ];
  if (q.t) lines.push('הערה: ' + esc_(q.t));
  const subject = (swap ? 'בקשת החלפה: ' + q.n + ' ⇄ ' + q.w : 'בקשה למסור משמרת: ' + q.n + ' → ' + q.w) + ' (' + q.m + ')';
  const html = '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;line-height:1.7">' +
    '<p style="font-size:17px;font-weight:bold;margin:0 0 8px">' + (swap ? 'בקשת החלפה חדשה' : 'בקשה חדשה למסור משמרת') + '</p>' +
    lines.map(function (l) { return '<div>' + l + '</div>'; }).join('') +
    '<p style="margin:18px 0"><a href="' + link + '" style="background:#1f6feb;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">פתח לאישור באתר השיבוץ</a></p>' +
    '<p style="color:#666;font-size:13px">האתר בודק את כל כללי השיבוץ לפני האישור ומראה אם ההחלפה יוצרת התנגשות. האישור עצמו נעשה באתר.</p></div>';
  const text = lines.join('\n').replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&') + '\n\nלאישור: ' + link;
  MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: subject, body: text, htmlBody: html, name: 'שיבוץ מתמחים' });
}

/* להרצה פעם אחת מהעורך: מאשרת את הרשאת שליחת המייל ושולחת מייל בדיקה */
function testMail() {
  notifyChg_({ m: '2026-11', n: 'בדיקה א', w: 'בדיקה ב', kind: 'swap', d1: 3, c1: 'n', d2: 5, c2: 'm', agree: true, t: 'זה מייל בדיקה' }, 2);
}


/* ---------- בדיקה אוטומטית של לוח התורנויות ---------- */
const BOARD_ID = '1IbVJDhWcut7F37GjC4nFQNwViArzh5nvR1EhfW4eoYQ';
const MIN_REST_H = 12, MAX_CONT_H = 26;
const ENG_M_ = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
// כותרת עמודה -> [שם, שעת התחלה, שעת סיום]; 'V' = תורנות צד ב (בסופ״ש 08–08)
const COLS_NEW_ = { 'בוקר 1-6': ['בוקר אגף א', 9, 21], 'לילה 1-6': ['לילה אגף א', 21, 9], 'אגף ג': ['אגף ג', 16, 8], 'פנימית': ['אגף ג', 16, 8], 'מהלכים פנימית': ['מהלכים פנימית', 16, 8], 'פנימית תורן שני': ['מהלכים פנימית', 16, 8],
  'כירורגיה בוקר': ['בוקר כירורגיה', 9, 21], 'כירורגיה ערב': ['ערב כירורגיה', 21, 9], 'טיפול נמרץ': ['טיפול נמרץ', 16, 8], 'אורתופדיה': ['אורתופדיה', 16, 8], 'ילדים': ['ילדים', 16, 8],
  'בוקר צד ב': ['בוקר צד ב', 8, 16], 'בוקר צד ג': ['בוקר צד ג', 8, 16], 'תורנות צד ב': 'V', 'מד״א': ['מד״א', 8, 16], 'בוקר אורתופדיה': ['בוקר אורתופדיה', 8, 16] };
const COLS_OLD_ = { 'בוקר א': ['בוקר אגף א', 9, 21], 'לילה א': ['לילה אגף א', 21, 9], 'בוקר ב': ['בוקר צד ב', 8, 16], 'בוקר כירורגיה': ['בוקר כירורגיה', 9, 21], 'ערב כירורגיה': ['ערב כירורגיה', 21, 9],
  'תורנות ב': 'V', 'פנימית': ['אגף ג', 16, 8], 'פנימית ב': ['מהלכים פנימית', 16, 8], 'אורטופדיה': ['אורתופדיה', 16, 8], 'ילדים': ['ילדים', 16, 8], 'טיפול נמרץ': ['טיפול נמרץ', 16, 8],
  'מד״א': ['מד״א', 8, 16], 'אורטו בוקר': ['בוקר אורתופדיה', 8, 16], 'אגף ג': ['בוקר צד ג', 8, 16] };

function nrm_(x) { return String(x === undefined || x === null ? '' : x).replace(/[‎‏‪-‮]/g, '').replace(/\s+/g, ' ').trim(); }

function boardShifts_(ss, y, m) {
  const sh = ss.getSheetByName(ENG_M_[m - 1] + ' ' + String(y).slice(2));
  if (!sh) return [];
  const vals = sh.getDataRange().getDisplayValues();
  const hi = vals.findIndex(function (r) { return r.some(function (c) { return nrm_(c) === 'יום'; }); });
  if (hi < 0) return [];
  const head = vals[hi].map(nrm_);
  const map = head.indexOf('בוקר א') >= 0 && head.indexOf('לילה א') >= 0 ? COLS_OLD_ : COLS_NEW_;
  const out = [];
  vals.slice(hi + 1).forEach(function (r) {
    const d = Number(nrm_(r[0]));
    if (!(d >= 1 && d <= 31)) return;
    const date = new Date(y, m - 1, d);
    if (date.getMonth() !== m - 1) return;
    const wk = date.getDay() === 5 || date.getDay() === 6;
    head.forEach(function (h, i) {
      let def = map[h];
      if (!def) return;
      if (def === 'V') def = wk ? ['בוקר + תורנות צד ב', 8, 8] : ['תורנות צד ב', 16, 8];
      nrm_(r[i]).split(/\s*[\/+,]\s*/).map(function (x) { return nrm_(x.replace(/\(.*?\)/g, '')); }).filter(String).forEach(function (name) {
        const s = new Date(y, m - 1, d, def[1]).getTime();
        let e = new Date(y, m - 1, d, def[2]).getTime();
        if (e <= s) e += 24 * 3600000;
        out.push({ name: name, d: d, m: m, y: y, what: def[0], s: s, e: e });
      });
    });
  });
  return out;
}

function boardConflicts_() {
  const ss = SpreadsheetApp.openById(BOARD_ID);
  const now = new Date(), y = now.getFullYear(), m = now.getMonth() + 1;
  const next = m === 12 ? [y + 1, 1] : [y, m + 1];
  const all = boardShifts_(ss, y, m).concat(boardShifts_(ss, next[0], next[1]));
  // בסופ״ש: מי שרשום גם בבוקר צד ב וגם בבוקר + תורנות צד ב, זו אותה משמרת
  const list = all.filter(function (a) { return !(a.what === 'בוקר צד ב' && all.some(function (b) { return b.name === a.name && b.d === a.d && b.m === a.m && b.what === 'בוקר + תורנות צד ב'; })); });
  const byName = {};
  list.forEach(function (x) { (byName[x.name] = byName[x.name] || []).push(x); });
  const today = new Date(y, m - 1, now.getDate()).getTime();
  const lbl = function (x) { return x.what + ' ' + x.d + '.' + x.m + ' (' + ('0' + new Date(x.s).getHours()).slice(-2) + ':00)'; };
  const issues = [];
  Object.keys(byName).forEach(function (n) {
    const L = byName[n].sort(function (a, b) { return a.s - b.s; });
    for (let i = 1; i < L.length; i++) {
      const a = L[i - 1], b = L[i], gap = (b.s - a.e) / 3600000;
      if (b.e < today) continue;
      const sameDay = a.d === b.d && a.m === b.m;
      if (gap < 0) issues.push({ t: b.s, text: n + ': ' + lbl(a) + ' חופף ל' + lbl(b) + ' — לא יכול להגיע' });
      else if (gap === 0 && sameDay) { if ((b.e - a.s) / 3600000 > MAX_CONT_H) issues.push({ t: b.s, text: n + ': ' + lbl(a) + ' ומיד ' + lbl(b) + ' — רצף מעל ' + MAX_CONT_H + ' שעות' }); }
      else if (gap < MIN_REST_H) issues.push({ t: b.s, text: n + ': ' + lbl(a) + ' ואחריו ' + lbl(b) + ' — מנוחה של ' + gap + ' שעות בלבד' + (gap === 0 ? ' (ישר מלילה/תורנות לבוקר)' : '') });
    }
  });
  return issues.sort(function (a, b) { return a.t - b.t; });
}

/* נקרא מהטריגרים: שולח מייל רק אם יש בעיה חדשה שעוד לא נשלחה */
function checkBoard() {
  const issues = boardConflicts_();
  try { quotaIssues_().forEach(function (x) { issues.push(x); }); } catch (err) { console.error('quota: ' + err); }
  const props = PropertiesService.getScriptProperties();
  const sent = JSON.parse(props.getProperty('boardSent') || '[]');
  const fresh = issues.filter(function (x) { return sent.indexOf(x.text) < 0; });
  if (fresh.length && NOTIFY_EMAIL) {
    const html = '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px;line-height:1.7"><p style="font-size:17px;font-weight:bold;margin:0 0 8px;color:#b42318">נמצאו בעיות בלוח התורנויות</p><ul>' +
      fresh.map(function (x) { return '<li>' + esc_(x.text) + '</li>'; }).join('') + '</ul>' +
      (issues.length > fresh.length ? '<p style="color:#666">ועוד ' + (issues.length - fresh.length) + ' בעיות שכבר נשלחו קודם ועדיין לא תוקנו.</p>' : '') +
      '<p><a href="https://docs.google.com/spreadsheets/d/' + BOARD_ID + '/edit">פתח את לוח התורנויות</a></p></div>';
    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: 'לוח התורנויות: ' + fresh.length + ' בעיות (מנוחה, חפיפה או חריגה ממכסה)', htmlBody: html, body: fresh.map(function (x) { return x.text; }).join('\n'), name: 'שיבוץ מתמחים' });
  }
  props.setProperty('boardSent', JSON.stringify(issues.map(function (x) { return x.text; })));
  return issues.map(function (x) { return x.text; });
}

/* להרצה פעם אחת מהעורך: בדיקה כל ערב ב־19:00, ובדיקה אחרי כל עריכה בלוח התורנויות */
function setupBoardCheck() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'checkBoard') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('checkBoard').timeBased().everyDays(1).atHour(19).inTimezone('Asia/Jerusalem').create();
  ScriptApp.newTrigger('checkBoard').forSpreadsheet(BOARD_ID).onEdit().create();
  PropertiesService.getScriptProperties().deleteProperty('boardSent');
  return checkBoard();
}


/* ---------- מסגרת אדומה סביב היום בלוח התורנויות ---------- */
function markToday() {
  const ss = SpreadsheetApp.openById(BOARD_ID);
  const props = PropertiesService.getScriptProperties();
  // מוריד את המסגרת הקודמת
  try {
    const prev = JSON.parse(props.getProperty('todayMark') || 'null');
    if (prev) {
      const ps = ss.getSheetByName(prev.sheet);
      if (ps) ps.getRange(prev.row, 1, 1, prev.width).setBorder(false, false, false, false, null, null);
    }
  } catch (err) {}
  const now = new Date();
  const tz = 'Asia/Jerusalem';
  const y = Number(Utilities.formatDate(now, tz, 'yyyy')), m = Number(Utilities.formatDate(now, tz, 'M')), d = Number(Utilities.formatDate(now, tz, 'd'));
  const sh = ss.getSheetByName(ENG_M_[m - 1] + ' ' + String(y).slice(2));
  if (!sh) { props.deleteProperty('todayMark'); return 'אין לשונית לחודש הזה'; }
  const vals = sh.getDataRange().getDisplayValues();
  const hi = vals.findIndex(function (r) { return r.some(function (c) { return nrm_(c) === 'יום'; }); });
  let row = -1;
  for (let i = hi + 1; i < vals.length; i++) if (Number(nrm_(vals[i][0])) === d) { row = i + 1; break; }
  if (row < 0) { props.deleteProperty('todayMark'); return 'לא נמצאה שורה ליום ' + d; }
  let width = vals[hi >= 0 ? hi : 0].length;
  while (width > 1 && !nrm_(vals[hi >= 0 ? hi : 0][width - 1])) width--;
  sh.getRange(row, 1, 1, width).setBorder(true, true, true, true, null, null, '#d93025', SpreadsheetApp.BorderStyle.SOLID_THICK);
  props.setProperty('todayMark', JSON.stringify({ sheet: sh.getName(), row: row, width: width }));
  return sh.getName() + ', שורה ' + row;
}

/* להרצה פעם אחת מהעורך: מסמן את היום עכשיו, ומעדכן כל לילה אחרי חצות */
function setupTodayMark() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'markToday') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('markToday').timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone('Asia/Jerusalem').create();
  return markToday();
}


/* להרצה פעם אחת מהעורך: מאשר את הגישה לסידור הכללי ובודק שהיא עובדת */
function testMainAccess() {
  const ss = SpreadsheetApp.openById(MAIN_ID);
  return ss.getName() + ': ' + ss.getSheets().length + ' לשוניות';
}


/* ---------- עמודות המתמחים בסידור הכללי, לפי לוח התורנויות ---------- */
const MAIN_FROM = '2026-11';
const HEB_M_ = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר'];
// כותרת בלוח התורנויות -> מספר עמודה בסידור הכללי
const MAIN_COLS_ = { 'בוקר 1-6': 5, 'לילה 1-6': 7, 'בוקר צד ב': 9, 'תורנות צד ב': 11, 'בוקר צד ג': 13, 'אגף ג': 15, 'פנימית': 15,
  'כירורגיה בוקר': 23, 'כירורגיה ערב': 24, 'מהלכים פנימית': 38, 'פנימית תורן שני': 38 };
const MAIN_SHARED_ = [15, 38]; // עמודות שגם מתמחי פנימית ממלאים: כותבים רק כשיש לנו שיבוץ
const SHORT_ = { 'מוחמד מיסק': 'מיסק', 'סימא שולמאן': 'סימא', 'באסל אבו עומר': 'באסל', 'דב מרטן': 'דב', 'אחמד עלי': 'עלי' };

function shortNames_(txt) { return nrm_(txt).split(/\s*\/\s*/).filter(String).map(function (x) { return SHORT_[x] || x; }).join(' / '); }

function syncMainMonth_(y, m) {
  const key = y + '-' + ('0' + m).slice(-2);
  if (key < MAIN_FROM) return 'לפני ' + MAIN_FROM;
  const board = SpreadsheetApp.openById(BOARD_ID).getSheetByName(ENG_M_[m - 1] + ' ' + String(y).slice(2));
  const main = SpreadsheetApp.openById(MAIN_ID).getSheetByName(HEB_M_[m - 1] + ' ' + y);
  if (!board || !main) return 'אין לשונית ' + (board ? 'בסידור הכללי' : 'בלוח התורנויות');
  const bv = board.getDataRange().getDisplayValues();
  const hi = bv.findIndex(function (r) { return r.some(function (c) { return nrm_(c) === 'יום'; }); });
  if (hi < 0) return 'אין כותרות';
  const head = bv[hi].map(nrm_);
  const want = {}, ours = {};
  bv.slice(hi + 1).forEach(function (r) {
    const d = Number(nrm_(r[0]));
    if (!(d >= 1 && d <= 31)) return;
    const wk = [5, 6].indexOf(new Date(y, m - 1, d).getDay()) >= 0;
    const row = {};
    head.forEach(function (h, i) {
      const c = MAIN_COLS_[h], v = nrm_(r[i]);
      if (!c || !v) return;
      row[c] = shortNames_(v);
      v.split(/\s*\/\s*/).forEach(function (x) { ours[nrm_(x)] = 1; ours[SHORT_[nrm_(x)] || nrm_(x)] = 1; });
    });
    if (wk && row[11] && !row[9]) row[9] = row[11];
    want[d] = row;
  });
  const mv = main.getDataRange().getDisplayValues();
  let n = 0;
  mv.forEach(function (r, i) {
    const d = Number(nrm_(r[0]));
    if (!(d >= 1 && d <= 31) || !want[d]) return;
    Object.keys(MAIN_COLS_).map(function (k) { return MAIN_COLS_[k]; }).filter(function (c, j, a) { return a.indexOf(c) === j; }).forEach(function (c) {
      const cur = nrm_(r[c - 1]), w = want[d][c] || '';
      if (cur === w) return;
      if (!w) {
        // אין לנו שיבוץ: מוחקים רק אם בתא כתוב אחד המתמחים שלנו
        const isOurs = cur.split(/\s*\/\s*/).every(function (x) { return ours[nrm_(x)]; });
        if (!cur || !isOurs) return;
      }
      main.getRange(i + 1, c).setValue(w);
      n++;
    });
  });
  return HEB_M_[m - 1] + ' ' + y + ': עודכנו ' + n + ' תאים';
}

/* מעדכן את החודש הנוכחי ואת החודש הבא */
function syncMain() {
  const now = new Date(), tz = 'Asia/Jerusalem';
  const y = Number(Utilities.formatDate(now, tz, 'yyyy')), m = Number(Utilities.formatDate(now, tz, 'M'));
  const out = [syncMainMonth_(y, m)];
  out.push(m === 12 ? syncMainMonth_(y + 1, 1) : syncMainMonth_(y, m + 1));
  return out.join(' | ');
}

/* להרצה פעם אחת מהעורך: עדכון אחרי כל עריכה בלוח התורנויות, ופעם ביום */
function setupMainSync() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncMain') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncMain').forSpreadsheet(BOARD_ID).onEdit().create();
  ScriptApp.newTrigger('syncMain').timeBased().everyDays(1).atHour(1).inTimezone('Asia/Jerusalem').create();
  return syncMain();
}

/* ---------- חריגה ממכסה: לפי לוח התורנויות והמכסות שבאתר השיבוץ ---------- */
const Q_SHIFT_ = { 'בוקר 1-6': 'm', 'לילה 1-6': 'n', 'כירורגיה בוקר': 'mc', 'כירורגיה ערב': 'nc', 'מד״א': 'a', 'בוקר אורתופדיה': 'z' };
const Q_DUTY_ = { 'תורנות צד ב': 'v', 'אגף ג': 'p', 'פנימית': 'p', 'מהלכים פנימית': 'c', 'פנימית תורן שני': 'c', 'טיפול נמרץ': 'i', 'אורתופדיה': 'o', 'ילדים': 'k' };
const DUTY_CODES_ = ['v', 'p', 'c', 'i', 'o', 'k'];
function nk_(x) { return nrm_(x).replace(/["'׳״`\-־.]/g, '').replace(/\s+/g, ''); }

function quotaMonth_(st, y, m) {
  const key = y + '-' + ('0' + m).slice(-2), mo = (st.months || {})[key];
  if (!mo) return [];
  const sh = SpreadsheetApp.openById(BOARD_ID).getSheetByName(ENG_M_[m - 1] + ' ' + String(y).slice(2));
  if (!sh) return [];
  const vals = sh.getDataRange().getDisplayValues();
  const hi = vals.findIndex(function (r) { return r.some(function (c) { return nrm_(c) === 'יום'; }); });
  if (hi < 0) return [];
  const head = vals[hi].map(nrm_), dim = new Date(y, m, 0).getDate(), stt = st.settings || {};
  const people = {}; (st.people || []).forEach(function (p) { people[p.id] = p; });
  const ids = Object.keys(mo.roster || {}).filter(function (id) { return people[id] && mo.roster[id].group !== 'X'; });
  const alias = st.nameAliases || {};
  const findId = function (n) {
    const k = nk_(n);
    let hit = ids.filter(function (id) { return nk_(people[id].name) === k; })[0];
    if (!hit && alias[k] && ids.indexOf(alias[k]) >= 0) hit = alias[k];
    if (!hit) {
      const t = nrm_(n).split(' ');
      const c = ids.filter(function (id) { const pt = nrm_(people[id].name).split(' '); return t.every(function (x) { return pt.indexOf(x) >= 0; }) || pt.every(function (x) { return t.indexOf(x) >= 0; }); });
      if (c.length === 1) hit = c[0];
    }
    return hit;
  };
  const cnt = {}; ids.forEach(function (id) { cnt[id] = { s: 0, d: 0, days: {} }; });
  vals.slice(hi + 1).forEach(function (r) {
    const d = Number(nrm_(r[0]));
    if (!(d >= 1 && d <= dim)) return;
    head.forEach(function (h, i) {
      const code = Q_SHIFT_[h] || Q_DUTY_[h];
      if (!code) return;
      nrm_(r[i]).split(/\s*[\/+,]\s*/).map(function (x) { return nrm_(x.replace(/\(.*?\)/g, '')); }).filter(String).forEach(function (n) {
        const id = findId(n); if (!id) return;
        (cnt[id].days[d] = cnt[id].days[d] || []).push(code);
      });
    });
  });
  const out = [], label = HEB_M_[m - 1] + ' ' + y;
  ids.forEach(function (id) {
    const r = mo.roster[id], c = cnt[id], g = r.group;
    Object.keys(c.days).forEach(function (d) {
      let L = c.days[d];
      if (L.indexOf('mc') >= 0 && L.indexOf('nc') >= 0) { c.d++; L = L.filter(function (x) { return x !== 'mc' && x !== 'nc'; }); }
      L.forEach(function (x) { if (DUTY_CODES_.indexOf(x) >= 0) c.d++; else if (g === 'A') c.s++; });
    });
    let vac = Number(r.extraVac) || 0;
    for (let d = 1; d <= dim; d++) { const cl = (mo.cells || {})[id + '|' + d]; if (cl && cl.mk === 'vac') vac++; }
    const frac = Math.max(0, 1 - vac / dim), ru = people[id].rules || {};
    let S = 0, D = 0;
    if (g === 'A') { S = (stt.baseAS || 11) * frac; D = (stt.baseAD || 3) * frac; }
    else D = ((ru.rotD !== undefined && ru.rotD !== '' && !isNaN(ru.rotD)) ? Number(ru.rotD) : (stt.baseRD || 6)) * frac;
    const course = stt.course || {};
    if (g === 'A' && (course.ids || []).indexOf(id) >= 0 && (course.dates || []).some(function (x) { return x.slice(0, 7) === key; })) S = Math.max(0, S - 1);
    if (r.tS !== null && r.tS !== '' && r.tS !== undefined && !isNaN(r.tS)) S = Number(r.tS);
    if (r.tD !== null && r.tD !== '' && r.tD !== undefined && !isNaN(r.tD)) D = Number(r.tD);
    S = Math.round(S); D = Math.round(D);
    const t = new Date(y, m - 1, 1).getTime();
    if (c.d > D) out.push({ t: t, text: people[id].name + ': ' + c.d + ' תורנויות ב' + label + ', המכסה ' + D + ' — חריגה ממכסה' });
    if (g === 'A' && c.s > S) out.push({ t: t, text: people[id].name + ': ' + c.s + ' משמרות ב' + label + ', המכסה ' + S + ' — חריגה ממכסה' });
  });
  return out;
}

function quotaIssues_() {
  const raw = readState_(); if (!raw) return [];
  const st = JSON.parse(raw);
  const now = new Date(), tz = 'Asia/Jerusalem';
  const y = Number(Utilities.formatDate(now, tz, 'yyyy')), m = Number(Utilities.formatDate(now, tz, 'M'));
  const nx = m === 12 ? [y + 1, 1] : [y, m + 1];
  return quotaMonth_(st, y, m).concat(quotaMonth_(st, nx[0], nx[1]));
}
