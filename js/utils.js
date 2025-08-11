// Utility helpers & shared formatting
/**
 * @typedef {Object} Employee
 * @property {string} id
 * @property {number} order
 * @property {string} group
 * @property {string} name
 * @property {string} position
 * @property {'waiter'|'hostess'|'fixed'|'hourly'} rateType
 * @property {number} [waiterPercent]
 * @property {number} [hostessPercent]
 * @property {number} [hourlyRate]
 * @property {number} [basePay]
 * @property {string} hoursText
 * @property {number} hoursMinutes
 * @property {number} sales
 * @property {number} gifts
 * @property {number} withheld
 * @property {number} pay
 */
export const $ = (sel) => document.querySelector(sel);
export const $$ = (sel) => Array.from(document.querySelectorAll(sel));
export const uid = () => Math.random().toString(36).slice(2, 11);

export function escapeHtml(str) {
  return String(str).replace(
    /[&<>"']/g,
    (s) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[s] || s
  );
}

export function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

export function extractNumber(str) {
  const m = String(str)
    .replace(/,/g, '.')
    .match(/(\d+[\.]?\d*)/);
  return m ? parseFloat(m[1]) : 0;
}

export function diffMinutes(a, b) {
  const [ah, am] = a.split(':').map(Number);
  const [bh, bm] = b.split(':').map(Number);
  let start = ah * 60 + am;
  let end = bh * 60 + bm;
  if (end < start) end += 1440;
  return end - start;
}

// Unified parse/validate of hours input. Returns { valid: boolean, minutes: number, token: string }
export function parseHoursInterval(raw) {
  const original = (raw || '').trim();
  if (!original) return { valid: true, minutes: 0, token: 'empty' };
  const clean = original.replace(/\s+/g, '');
  // Day-off / absence tokens (UA/RU variants)
  if (/^(в|вихід\w*|вибув\w*|виход\w*|выход\w*|выбыл\w*)$/i.test(clean)) {
    return { valid: true, minutes: 0, token: 'dayoff' };
  }
  const m = clean.match(/^(\d{1,2}:\d{2})-(\d{1,2}:\d{2})$/);
  if (!m) return { valid: false, minutes: 0, token: 'pattern' };
  const validTime = (t, isEnd = false) => {
    const [h, mi] = t.split(':').map(Number);
    if (isNaN(h) || isNaN(mi)) return false;
    if (h < 0 || h > 24) return false;
    if (mi < 0 || mi > 59) return false;
    if (h === 24 && mi !== 0) return false; // only 24:00 allowed
    if (!isEnd && h === 24) return false; // start cannot be 24:00
    return true;
  };
  if (!(validTime(m[1]) && validTime(m[2], true)))
    return { valid: false, minutes: 0, token: 'range' };
  const minutes = diffMinutes(m[1], m[2] === '24:00' ? '00:00' : m[2]);
  return { valid: true, minutes, token: 'interval' };
}

/**
 * Human readable display of rate column (unified across UI / export)
 * @param {Employee} emp
 */
export function rateDisplay(emp) {
  if (!emp) return '';
  if (emp.rateType === 'waiter') return (emp.waiterPercent || 5) + '%';
  if (emp.rateType === 'hostess') return String(emp.hourlyRate || 0);
  if (emp.rateType === 'fixed') return String(emp.basePay || 0);
  return String(emp.hourlyRate || 0);
}
