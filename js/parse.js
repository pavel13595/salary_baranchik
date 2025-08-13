import { uid, extractNumber, diffMinutes } from './utils.js';
import { state } from './state.js';

export function interpretRate(position, rateStr) {
  const lowerPos = (position || '').toLowerCase();
  const lower = (rateStr || '').toLowerCase();
  if (/(официант|офіціант)/i.test(position) || /5\s*%/.test(lower)) {
    return { rateType: 'waiter', waiterPercent: 5, hostessPercent: 0, hourlyRate: 0, basePay: 0 };
  }
  if (/хостес/i.test(position)) {
    const hr = extractNumber(rateStr);
    return { rateType: 'hostess', hourlyRate: hr, hostessPercent: 2, waiterPercent: 0, basePay: 0 };
  }
  if (/фікс|фикс/.test(lower)) {
    const num = extractNumber(lower);
    return { rateType: 'fixed', hourlyRate: 0, waiterPercent: 0, hostessPercent: 0, basePay: num };
  }
  const hr = extractNumber(lower);
  return { rateType: 'hourly', hourlyRate: hr, waiterPercent: 0, hostessPercent: 0, basePay: 0 };
}

export function parseEmployeesBulk(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trimEnd());
  let order = 1;
  let currentGroup = '';
  const employees = [];
  const knownPositions = [
    'официант',
    'офіціант',
    'официантка',
    'офіціантка',
    'хостес',
    'бармен',
    "кур'єр",
    'курєр',
    'курьер',
    'пакувальниця',
    'пакувальник',
    'господиня',
    'завгосп',
    'завхоз',
    'менеджер',
    'керуючий',
    'су-шеф',
    'су шеф',
    'кухар',
    'повар',
    'ранер',
    'раннер',
  ];
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const posRe = new RegExp(knownPositions.map(esc).join('|'), 'i');
  const RATE_NUM_TAIL_RE = /(\d+[\.,]?\d*)\s*(?:грн|uah|₴)?(?:\s*[\\\/]\s*(?:год|час|ч))?\s*$/i;
  function detectTrailingRate(str) {
    if (!str) return null;
    let m = str.match(/(\d+\s*%)\s*$/);
    if (m) return { rateStr: m[1], stripped: str.slice(0, m.index).trim() };
    m = str.match(/((?:фікс|фикс|fix)\s*\d+[\.,]?\d*)\s*$/i);
    if (m) return { rateStr: m[1], stripped: str.slice(0, m.index).trim() };
    m = str.match(RATE_NUM_TAIL_RE);
    if (m) return { rateStr: m[0].trim(), stripped: str.slice(0, m.index).trim() };
    return null;
  }
  for (let raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    // 1) Primary: split by tabs
    let parts = raw
      .split(/\t+/)
      .map((s) => s.trim())
      .filter(Boolean);
    // 2) Fallback: split by common separators (two+ spaces, comma, semicolon, pipe, dash surrounded by spaces)
    if (parts.length < 2) {
      parts = raw
        .split(/\s{2,}|\s+[-–—]\s+|\s*\|\s*|\s*;\s*|\s*,\s*/)
        .map((s) => s.trim())
        .filter(Boolean);
    }
    // 3) If still not enough parts, try heuristic parse using known position words and trailing rate
    let name, position, rateStr;
    if (parts.length >= 2) {
      if (parts.length > 3) {
        rateStr = parts.pop();
        position = parts.pop();
        name = parts.join(' ');
        // Prefer embedded rate in position if last token is empty/0 or not a rate
        const emb = detectTrailingRate(position);
        const lastVal = extractNumber(rateStr || '');
        if (emb && (!rateStr || isNaN(lastVal) || lastVal === 0)) {
          rateStr = emb.rateStr;
          position = emb.stripped;
        }
      } else if (parts.length === 3) {
        [name, position, rateStr] = parts;
        // If position contains embedded rate and the provided rate is 0/invalid, use embedded
        const emb = detectTrailingRate(position);
        const lastVal = extractNumber(rateStr || '');
        if (emb && (rateStr === undefined || isNaN(lastVal) || lastVal === 0)) {
          rateStr = emb.rateStr;
          position = emb.stripped;
        }
      } else {
        // exactly two parts
        const p1 = parts[0];
        const p2 = parts[1];
        // If second token looks like a rate, treat it as rate and split p1 into name/position
        const looksRate =
          /(\d+\s*%$)|((?:фікс|фикс|fix)\s*\d+[\.,]?\d*$)|((?:\d+[\.,]?\d*)\s*(?:грн|uah|₴)?(?:\s*[\\\/]\s*(?:год|час|ч))?$)/i.test(
            p2
          );
        if (looksRate) {
          rateStr = p2;
          // Try find a known position inside p1
          const mP = p1.match(posRe);
          if (mP) {
            position = p1.slice(mP.index).trim();
            name = p1.slice(0, mP.index).trim();
          } else {
            const lastSpace = p1.lastIndexOf(' ');
            if (lastSpace > 0) {
              name = p1.slice(0, lastSpace).trim();
              position = p1.slice(lastSpace + 1).trim();
            } else {
              name = p1.trim();
              position = '';
            }
          }
        } else {
          // Otherwise assume [name, position]
          name = p1;
          position = p2;
          rateStr = '';
          // Also check if position has embedded rate
          const emb = detectTrailingRate(position);
          if (emb) {
            rateStr = emb.rateStr;
            position = emb.stripped;
          }
        }
      }
    } else {
      // Heuristic: try to extract rate at the end (percent or number with optional currency)
      let tmp = line;
      const emb = detectTrailingRate(tmp);
      if (emb) {
        rateStr = emb.rateStr;
        tmp = emb.stripped;
      }
      // Try find position keyword inside remaining
      const mPos = tmp.match(posRe);
      if (mPos) {
        const before = tmp.slice(0, mPos.index).trim();
        // If position keyword is at start and there's no rate and no name prefix,
        // it's likely a group header like "Хостес/Доставка" -> treat as group
        if (!rateStr && before.length === 0) {
          currentGroup = tmp;
          continue;
        }
        position = tmp.slice(mPos.index).trim();
        name = before;
      } else {
        // As a last resort, consider it's a group line if no rate detected
        if (!rateStr) {
          currentGroup = tmp;
          continue;
        }
        // Otherwise split by last space into name/position
        const lastSpace = tmp.lastIndexOf(' ');
        if (lastSpace > 0) {
          name = tmp.slice(0, lastSpace).trim();
          position = tmp.slice(lastSpace + 1).trim();
        } else {
          name = tmp;
          position = '';
        }
      }
    }
    const parsed = interpretRate(position, rateStr || '');
    employees.push({
      id: uid(),
      order: order++,
      group: currentGroup,
      name: name || '',
      position: position || '',
      rawRateStr: rateStr || '',
      ...parsed,
      hoursText: '',
      hoursMinutes: 0,
      sales: 0,
      gifts: 0,
      withheld: 0,
      pay: 0,
    });
  }
  return employees;
}

export function parseHoursBulk(text) {
  return text.split(/\r?\n/).map((l) => l.trim());
}

export function applyHoursToEmployees(hoursLines) {
  const lines = hoursLines.map((l) => l.trim());
  let i = 0;
  const n = lines.length;
  for (const emp of state.employees) {
    if (emp.rateType === 'fixed') {
      emp.hoursText = '';
      emp.hoursMinutes = 0;
      continue;
    }
    while (i < n && lines[i] === '') i++;
    if (i >= n) {
      emp.hoursText = '';
      emp.hoursMinutes = 0;
      continue;
    }
    let token = lines[i];
    i++;
    if (!token) {
      emp.hoursText = '';
      emp.hoursMinutes = 0;
      continue;
    }
    if (/^в$/i.test(token)) {
      emp.hoursText = 'в';
      emp.hoursMinutes = 0;
      continue;
    }
    if (/вихід|вихідн|вибув/i.test(token)) {
      emp.hoursText = token;
      emp.hoursMinutes = 0;
      continue;
    }
    token = token.replace(/\s+/g, '');
    const m = token.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
    if (m) {
      const minutes = diffMinutes(m[1], m[2]);
      emp.hoursText = `${m[1]}-${m[2]}`;
      emp.hoursMinutes = minutes;
    } else {
      emp.hoursText = token;
      emp.hoursMinutes = 0;
    }
  }
}
