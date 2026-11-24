/* משותף לטופס הבכירים ולאתר הניהול */
const API_DEFAULT = 'https://script.google.com/macros/s/AKfycbzWke13nl2Kt9ctKqor5dL2UHL7JY5bZrS6UbWQoKTPw5M4ZhMwLufHe_U4CPykRoBB/exec';
const HEB_MONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
const WD = ['א','ב','ג','ד','ה','ו','ש'];
const AREAS = { K: 'כונן', A: 'שוכבים א', B: 'שוכבים ב', C: 'שוכבים ג', O: 'אגף ד', W: 'מהלכים', S: 'שוכבים ערב', N: 'לילה' };
/* תואם לעמודות הבכירים בסידור הכללי של המלר״ד (col = מספר העמודה בגיליון, A=1).
   opt = משבצת רשות: לא נספרת כחסר. oc = כוננות: לא מתנגשת בשעות. days: 0=א … 6=ש */
const DEFAULT_SLOTS = [
  { id: 'K',  area: 'K', name: 'כונן',            from: 0,  to: 24, oc: true, t: 'כוננות', col: 3,  days: [0,1,2,3,4,5,6] },
  { id: 'A1', area: 'A', name: 'שוכבים א',           from: 8,  to: 16, col: 4,  days: [0,1,2,3,4,5] },
  { id: 'B1', area: 'B', name: 'שוכבים ב',           from: 8,  to: 16, col: 8,  days: [0,1,2,3,4] },
  { id: 'C1', area: 'C', name: 'שוכבים ג',           from: 8,  to: 16, col: 12, days: [0,1,2,3,4,5] },
  { id: 'O1', area: 'O', name: 'אגף ד',           from: 8,  to: 16, col: 16, days: [0,1,2,3,4], opt: true },
  { id: 'W1', area: 'W', name: 'מהלכים צוות א',   from: 8,  to: 16, col: 18, days: [0,1,2,3,4,5] },
  { id: 'W3', area: 'W', name: 'מהלכים צוות ב',   from: 10, to: 17, col: 20, days: [0,1,2,3,4] },
  { id: 'W4', area: 'W', name: 'מהלכים צוות ג',   from: 11, to: 18, col: 21, days: [0,1,2,3,4] },
  { id: 'E1', area: 'W', name: 'ערב מהלכים 1',    from: 16, to: 24, col: 31, days: [0,1,2,3,4] },
  { id: 'E2', area: 'W', name: 'ערב מהלכים 2',    from: 16, to: 24, col: 32, days: [0,1,2,3,4] },
  { id: 'E3', area: 'S', name: 'ערב שוכבים 1',    from: 16, to: 24, col: 33, days: [0,1,2,3,4] },
  { id: 'E4', area: 'S', name: 'ערב שוכבים 2',    from: 16, to: 24, col: 34, days: [0,1,2,3,4], opt: true },
  { id: 'N1', area: 'N', name: 'לילה',            from: 24, to: 32, t: 'לילה', col: 36, days: [0,1,2,3,4,5,6], opt: true }
];
const MAIN_ID = '1KCavu2KRgD4DzspCwB391_PZmdQiXh3-Sj3H0HKnVHo';
const HOLIDAYS = { Tishri: { 1: ['hol','ראש השנה'], 2: ['hol','ראש השנה'], 9: ['eve','ערב יום כיפור'], 10: ['hol','יום כיפור'], 14: ['eve','ערב סוכות'], 15: ['hol','סוכות'], 21: ['eve','ערב שמחת תורה'], 22: ['hol','שמיני עצרת ושמחת תורה'] },
  Elul: { 29: ['eve','ערב ראש השנה'] }, Nisan: { 14: ['eve','ערב פסח'], 15: ['hol','פסח'], 20: ['eve','ערב שביעי של פסח'], 21: ['hol','שביעי של פסח'] }, Sivan: { 5: ['eve','ערב שבועות'], 6: ['hol','שבועות'] } };
/* חלקי יום לבקשות: '' כל היום, m בוקר (משמרת שמתחילה לפני 16), e ערב (משמרת שנגמרת אחרי 16) */
const PARTS = { '': 'כל היום', m: 'בוקר', e: 'ערב' };

const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad2 = n => String(n).padStart(2, '0');
const norm = s => String(s || '').replace(/[‎‏‪-‮⁦-⁩﻿­]/g, '').trim().replace(/\s+/g, ' ');
function slotTimeTxt(s) { return s.t ? s.t : `${pad2(s.from)}–${pad2(s.to)}`; }
function slotTime(s) { return s.t ? s.t : '\u2066' + slotTimeTxt(s) + '\u2069'; }
function slotName(s) { return s.name || AREAS[s.area] || s.area; }
function slotLabel(s) { return s.t ? slotName(s) : `${slotName(s)} ${slotTime(s)}`; }
function slotShort(s) { return slotLabel(s); }
function overlap(a, b) { if (a.oc || b.oc) return false; return a.from < b.to && b.from < a.to; }
function partHits(part, s) { return !part || !!s.oc || (part === 'm' ? s.from < 16 : s.to > 16); }
function slotDays(s) { return Array.isArray(s.days) ? s.days : [0,1,2,3,4,5,6]; }

const _hc = {};
function kDim(key) { const [y, m] = key.split('-').map(Number); return new Date(y, m, 0).getDate(); }
function kWd(key, d) { const [y, m] = key.split('-').map(Number); return new Date(y, m - 1, d).getDay(); }
function hol(key) {
  if (_hc[key]) return _hc[key];
  const out = {}, [y, m] = key.split('-').map(Number);
  try {
    const f = new Intl.DateTimeFormat('en-u-ca-hebrew', { day: 'numeric', month: 'long' });
    for (let d = 1; d <= kDim(key); d++) {
      const p = f.formatToParts(new Date(y, m - 1, d));
      const hm = p.find(x => x.type === 'month')?.value, hd = Number(p.find(x => x.type === 'day')?.value);
      const h = HOLIDAYS[hm] && HOLIDAYS[hm][hd];
      if (h) out[d] = { t: h[0], name: h[1] };
    }
  } catch (e) {}
  return (_hc[key] = out);
}
function kWeekend(key, d) { const w = kWd(key, d); return w === 5 || w === 6 || hol(key)[d]?.t === 'hol'; }
function dayClass(key, d) { const h = hol(key)[d]; return h ? h.t : (kWd(key, d) === 5 || kWd(key, d) === 6 ? 'wk' : ''); }
function kLabel(key) { const [y, m] = key.split('-').map(Number); return `${HEB_MONTHS[m - 1]} ${y}`; }
function monthOpts(back = 0, fwd = 3) { const n = new Date(), o = []; for (let i = -back; i < fwd; i++) { const dt = new Date(n.getFullYear(), n.getMonth() + i, 1); o.push(`${dt.getFullYear()}-${pad2(dt.getMonth() + 1)}`); } return o; }
function inVac(vac, d) { return (vac || []).some(([a, b]) => d >= a && d <= b); }

function toast(msg, ms = 2800) { const el = document.createElement('div'); el.className = 'toast'; el.textContent = msg; document.body.appendChild(el); setTimeout(() => el.remove(), ms); }
function openSheet(html) {
  const L = $('#layer'); const back = document.createElement('div'); back.className = 'back';
  back.innerHTML = `<div class="sheet" role="dialog" aria-modal="true">${html}</div>`; L.appendChild(back);
  const close = () => back.remove();
  back.addEventListener('click', e => { if (e.target === back || e.target.closest('[data-close]')) close(); });
  return { el: back.querySelector('.sheet'), close };
}
function copyText(t) { (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(() => toast('הועתק'), () => { const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); toast('הועתק'); } catch (e) {} ta.remove(); }); }
function getApi() {
  let api = API_DEFAULT.startsWith('http') ? API_DEFAULT : '';
  try {
    const u = new URLSearchParams(location.search).get('api');
    if (u && /^https:\/\/script\.google(usercontent)?\.com\//.test(u)) { api = u; localStorage.setItem('sr-api', u); }
    else if (!api) api = localStorage.getItem('sr-api') || '';
  } catch (e) {}
  return api;
}
async function apiGet(api, params) { const r = await fetch(api + '?' + new URLSearchParams(params)); return r.json(); }
async function apiPost(api, body) { const r = await fetch(api, { method: 'POST', body: JSON.stringify(body) }); return r.json(); }
function lsGet(k, def) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch (e) { return def; } }
function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
