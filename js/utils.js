// Utility helpers
export const $ = sel => document.querySelector(sel);
export const $$ = sel => Array.from(document.querySelectorAll(sel));
export const uid = () => Math.random().toString(36).slice(2, 11);

export function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, s => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[s] || s));
}

export function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

export function extractNumber(str) {
  const m = String(str).replace(/,/g, '.').match(/(\d+[\.]?\d*)/);
  return m ? parseFloat(m[1]) : 0;
}

export function diffMinutes(a, b) {
  const [ah, am] = a.split(':').map(Number); const [bh, bm] = b.split(':').map(Number);
  let start = ah * 60 + am; let end = bh * 60 + bm; if (end < start) end += 1440; return end - start;
}
