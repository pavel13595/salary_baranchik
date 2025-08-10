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
  const lines = text.split(/\r?\n/).map(l => l.trimEnd());
  let order = 1;
  let currentGroup = '';
  const employees = [];
  for (let raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    const parts = raw.split(/\t+/).map(s => s.trim()).filter(Boolean);
    if (parts.length === 1) { currentGroup = parts[0]; continue; }
    let name, position, rateStr;
    if (parts.length > 3) {
      rateStr = parts.pop();
      position = parts.pop();
      name = parts.join(' ');
    } else {
      [name, position, rateStr] = parts;
    }
    const parsed = interpretRate(position, rateStr || '');
    employees.push({
      id: uid(), order: order++, group: currentGroup, name: name || '', position: position || '', rawRateStr: rateStr || '',
      ...parsed, hoursText: '', hoursMinutes: 0, sales: 0, gifts: 0, withheld: 0, pay: 0
    });
  }
  return employees;
}

export function parseHoursBulk(text) { return text.split(/\r?\n/).map(l => l.trim()); }

export function applyHoursToEmployees(hoursLines) {
  const lines = hoursLines.map(l => l.trim());
  let i = 0; const n = lines.length;
  for (const emp of state.employees) {
    if (emp.rateType === 'fixed') { emp.hoursText = ''; emp.hoursMinutes = 0; continue; }
    while (i < n && lines[i] === '') i++;
    if (i >= n) { emp.hoursText = ''; emp.hoursMinutes = 0; continue; }
    let token = lines[i]; i++;
    if (!token) { emp.hoursText = ''; emp.hoursMinutes = 0; continue; }
    if (/^в$/i.test(token)) { emp.hoursText = 'в'; emp.hoursMinutes = 0; continue; }
    if (/вихід|вихідн|вибув/i.test(token)) { emp.hoursText = token; emp.hoursMinutes = 0; continue; }
    token = token.replace(/\s+/g, '');
    const m = token.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
    if (m) { const minutes = diffMinutes(m[1], m[2]); emp.hoursText = `${m[1]}-${m[2]}`; emp.hoursMinutes = minutes; }
    else { emp.hoursText = token; emp.hoursMinutes = 0; }
  }
}
