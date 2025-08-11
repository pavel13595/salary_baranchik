import { state } from './state.js';
import { showToast } from './ui.js';
import { computePays } from './pay.js';
import { rateDisplay } from './utils.js';

const FIXED_LAYOUT = [
  { key: 'name', title: '', width: 28 },
  { key: 'position', title: 'Посада', width: 16 },
  { key: 'hours', title: 'Кількість відпрацьованих відпрацьованих годин, %, змін', width: 18 },
  { key: 'rate', title: 'Ставка', width: 10 },
  { key: 'sales', title: 'Продажі', width: 12 },
  { key: 'withheld', title: 'Утримано', width: 12 },
  { key: 'issued', title: 'Видано', width: 12 },
  { key: 'total', title: 'Всього нараховано', width: 18 },
  { key: 'sign', title: 'Підпис отримувача', width: 18 }
];

const SUBGROUPS = [
  { name: 'Адмін. Персонал', patterns: [/^керуюч(ий|а)$/i, /^менеджер$/i, /^шеф$/i] },
  { name: 'Бар', patterns: [/бармен/i, /бармени/i] },
  { name: 'Кухня', patterns: [/кухар/i, /су[- ]?шеф/i, /суш[- ]?шеф/i] },
  { name: 'Офіціанти / ранери', patterns: [/офіціант/i, /официант/i, /ранер/i] },
  { name: 'Хостес / Доставка', patterns: [/хостес/i, /пакувальниц/i, /пакувальник/i, /упаковщ/i, /упаковщик/i] },
  { name: 'Господарка', patterns: [/господарк/i, /господин/i] },
  { name: 'Інший персонал', patterns: [/кур[’'`]?єр/i, /курьер/i, /завгосп/i] }
];
function classifySubgroup(position) { const pos = (position || '').toLowerCase(); for (const sg of SUBGROUPS) { if (sg.patterns.some(rx => rx.test(pos))) return sg.name; } return 'Інший персонал'; }

function buildReportData() {
  const repDateStr = state.settings.reportDate || (() => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().slice(0, 10); })();
  const dParts = repDateStr.split('-');
  const dateStrDisp = `${dParts[2]}.${dParts[1]}`;
  const titleLine = `Той Самий Баранчик${state.settings.city ? ' ' + state.settings.city : ''}`;
  const headersLocal = FIXED_LAYOUT.map(c => c.title);
  const data = []; const merges = [];
  data.push([titleLine, ...Array(headersLocal.length - 1).fill('')]); merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: headersLocal.length - 1 } });
  const dateRow = [dateStrDisp, ...Array(headersLocal.length - 1).fill('')];
  data.push(dateRow); merges.push({ s: { r: 1, c: 1 }, e: { r: 1, c: headersLocal.length - 1 } });
  data.push(headersLocal);
  let rowIdx = 3;
  const isDayOff = emp => /вихід|вибув/i.test(emp.hoursText || '');
  const totalColIndex = FIXED_LAYOUT.findIndex(c => c.key === 'total');
  const subgroupMap = new Map(); SUBGROUPS.forEach(sg => subgroupMap.set(sg.name, []));
  for (const emp of state.employees) { if (isDayOff(emp)) continue; if (!(emp.rateType === 'fixed' || emp.hoursMinutes > 0)) continue; const sg = classifySubgroup(emp.position); if (!subgroupMap.has(sg)) subgroupMap.set(sg, []); subgroupMap.get(sg).push(emp); }
  function rateDisp(e) { return rateDisplay(e); }
  function hoursF(e) { return e.hoursMinutes ? (e.hoursMinutes / 60).toFixed(2) : (e.hoursMinutes === 0 ? '0' : ''); }
  let totalAll = 0;
  for (const sg of SUBGROUPS.map(s => s.name)) {
    const emps = subgroupMap.get(sg) || []; if (!emps.length) continue;
    data.push([sg, ...Array(headersLocal.length - 1).fill('')]); merges.push({ s: { r: rowIdx, c: 0 }, e: { r: rowIdx, c: headersLocal.length - 1 } }); rowIdx++;
    let groupSum = 0;
    for (const e of emps) { const payInt = (typeof e.pay === 'number') ? Math.round(e.pay) : ''; const row = []; row.push(e.name); row.push(e.position); row.push(hoursF(e)); row.push(rateDisp(e)); row.push((e.rateType === 'waiter' || e.rateType === 'hostess') ? (e.sales || 0) : ''); row.push(e.withheld || ''); row.push(''); row.push(payInt); groupSum += (typeof payInt === 'number' ? payInt : 0); totalAll += (typeof payInt === 'number' ? payInt : 0); row.push(''); data.push(row); rowIdx++; }
    const subtotalLabel = `всього ${sg.toLowerCase()}`; const subtotalRow = Array(headersLocal.length).fill(''); subtotalRow[0] = subtotalLabel; subtotalRow[totalColIndex] = groupSum; data.push(subtotalRow); const issuedIdx = FIXED_LAYOUT.findIndex(c => c.key === 'issued'); merges.push({ s: { r: rowIdx, c: 1 }, e: { r: rowIdx, c: issuedIdx } }); rowIdx++;
  }
  const grand = Array(headersLocal.length).fill(''); grand[0] = 'ВСЬОГО'; grand[totalColIndex] = totalAll; data.push(grand); const issuedIdx = FIXED_LAYOUT.findIndex(c => c.key === 'issued'); merges.push({ s: { r: rowIdx, c: 1 }, e: { r: rowIdx, c: issuedIdx } });
  return { headers: headersLocal, data, merges, repDateStr };
}

export async function exportExcel() {
  // Safety: ensure latest calculations (in case something changed without re-render)
  computePays();
  // Block export if any invalid hours exist (data integrity safeguard)
  try {
    const hoursInvalid = state.employees.some(emp => {
      const raw = (emp.hoursText || '').trim();
      if (!raw) return false; // empty allowed
      const clean = raw.replace(/\s+/g, '');
      // Allowed tokens for absence / special status
      if (/^(в|вихід|вихідн|вибув)$/i.test(clean)) return false;
      const m = clean.match(/^(\d{1,2}:\d{2})-(\d{1,2}:\d{2})$/);
      if (!m) return true; // pattern mismatch
      const validTime = t => { const [h, mi] = t.split(':').map(Number); return h >= 0 && h < 24 && mi >= 0 && mi < 60; };
      if (!(validTime(m[1]) && validTime(m[2]))) return true;
      return false;
    });
    if (hoursInvalid) { showToast('Є некоректно введений час — виправте перед експортом', 'error', 6000); return; }
  } catch (e) { /* fail-safe: if validation throws, continue with export attempt */ }
  if (window.ExcelJS) { return exportExcelExcelJS(); }
  if (!state.employees.length) { showToast('Немає даних', 'error'); return; }
  if (!window.XLSX) { showToast('XLSX не завантажено', 'error'); return; }
  const { headers, data, merges, repDateStr } = buildReportData();
  try {
    const ws = XLSX.utils.aoa_to_sheet(data);
    if (merges.length) ws['!merges'] = merges;
    ws['!cols'] = FIXED_LAYOUT.map(c => ({ wch: c.width }));
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.length - 1, c: headers.length - 1 } });
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Відомість');
    applyExactVisualTemplateSheetJS(ws);
    const citySegment = state.settings.city ? state.settings.city.replace(/\s+/g, '_') : 'Місто';
    const dParts = repDateStr.split('-');
    const dateSeg = `${dParts[2]}.${dParts[1]}`;
    const filename = `Відомість_${citySegment}_${dateSeg}.xlsx`;
    XLSX.writeFile(wb, filename, { compression: true });
    showToast(`Експортовано (обмежений стиль) ${filename}`, 'success');
  } catch (err) { console.error(err); showToast('Помилка експорту', 'error'); }
}

async function exportExcelExcelJS() {
  computePays();
  // Duplicate guard (in case exportExcelExcelJS invoked directly or by future refactor)
  try {
    const hoursInvalid = state.employees.some(emp => {
      const raw = (emp.hoursText || '').trim();
      if (!raw) return false;
      const clean = raw.replace(/\s+/g, '');
      if (/^(в|вихід|вихідн|вибув)$/i.test(clean)) return false;
      const m = clean.match(/^(\d{1,2}:\d{2})-(\d{1,2}:\d{2})$/);
      if (!m) return true;
      const validTime = t => { const [h, mi] = t.split(':').map(Number); return h >= 0 && h < 24 && mi >= 0 && mi < 60; };
      if (!(validTime(m[1]) && validTime(m[2]))) return true;
      return false;
    });
    if (hoursInvalid) { showToast('Є некоректно введений час — виправте перед експортом', 'error', 6000); return; }
  } catch (_) { }
  if (!state.employees.length) { showToast('Немає даних', 'error'); return; }
  const { data, merges, repDateStr } = buildReportData();
  const citySegment = state.settings.city ? state.settings.city.replace(/\s+/g, '_') : 'Місто';
  const dParts = repDateStr.split('-'); const dateSeg = `${dParts[2]}.${dParts[1]}`; const filename = `Відомість_${citySegment}_${dateSeg}.xlsx`;
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Відомість', { properties: { defaultRowHeight: 12.75 } });
    data.forEach(r => ws.addRow(r));
    function colLetter(idx) { let s = ''; idx = idx + 1; while (idx > 0) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = (idx - 1) / 26 | 0; } return s; }
    const lastColIndex = FIXED_LAYOUT.length - 1; const lastColLetter = colLetter(lastColIndex);
    const dynMerges = merges.map(m => `${colLetter(m.s.c)}${m.s.r + 1}:${colLetter(m.e.c)}${m.e.r + 1}`);
    const STATIC_MERGES = ['A1:' + lastColLetter + '1', 'B2:' + lastColLetter + '2'];
    const maxRow = ws.rowCount; function mergeSafe(range) { try { const rowNums = range.match(/\d+/g); if (rowNums) { const maxR = Math.max(...rowNums.map(Number)); if (maxR > maxRow) return; } ws.mergeCells(range); } catch (e) { } }
    [...dynMerges, ...STATIC_MERGES].forEach(mergeSafe);
    const setW = (col, val) => { if (ws.getColumn(col)) ws.getColumn(col).width = val; };
    setW('A', 30.7109375); setW('B', 19.5703125); setW('C', 9.28515625); setW('G', 11.5703125); setW('H', 12.140625); setW('I', 14.7109375); setW('J', 14.42578125); setW('K', 14.42578125);
    ['D', 'E', 'F'].forEach(c => { if (ws.getColumn(c)) ws.getColumn(c).width = 9.3; });
    const keyHeights = { 1: 12.75, 2: 12.75, 3: 65, 4: 12.75 }; Object.entries(keyHeights).forEach(([r, h]) => { if (ws.getRow(+r)) ws.getRow(+r).height = h; });
    for (let r = 5; r <= ws.rowCount; r++) { if (!ws.getRow(r).height) ws.getRow(r).height = 12.75; }
    const FONT = { name: 'Calibri', size: 11, color: { argb: 'FF000000' } };
    const alignC = { horizontal: 'center', vertical: 'middle', wrapText: true };
    const alignNum = { horizontal: 'center', vertical: 'middle' };
    const thin = { style: 'thin', color: { argb: 'FF000000' } };
    const medium = { style: 'medium', color: { argb: 'FF000000' } };
    let headerRowIndex = 3;
    for (let r = 1; r <= ws.rowCount; r++) {
      const rowObj = ws.getRow(r);
      for (let c = 1; c <= rowObj.cellCount; c++) {
        const v = String(ws.getCell(r, c).value || '').trim();
        if (v === 'Посада') { headerRowIndex = r; r = ws.rowCount + 1; break; }
      }
    }
    ws.getCell('A1').font = { ...FONT, bold: true }; ws.getCell('A1').alignment = alignC;
    ws.getRow(1).eachCell(cell => { cell.border = { ...(cell.border || {}), bottom: medium, top: cell.border?.top }; });
    if (ws.getRow(2)) { ws.getRow(2).eachCell(c => { c.font = { ...FONT, bold: true }; c.alignment = alignC; }); }
    if (headerRowIndex) { const hr = ws.getRow(headerRowIndex); hr.eachCell(cell => { cell.font = { ...FONT, bold: true }; cell.alignment = alignC; cell.border = { top: medium, bottom: medium, left: thin, right: thin }; }); }
    for (let r = 1; r <= ws.rowCount; r++) {
      if (r === 1 || r === 2 || r === headerRowIndex) continue;
      const row = ws.getRow(r); const valA = String(ws.getCell(`A${r}`).value || '').trim(); const cellB = ws.getCell(`B${r}`); const cellC = ws.getCell(`C${r}`);
      const isGrandTotal = /^всього$/i.test(valA) || /^ВСЬОГО$/.test(valA);
      const isGroupTotal = /^всього\s+/.test(valA.toLowerCase()) && !isGrandTotal;
      const isSubgroup = !isGrandTotal && !isGroupTotal && !!valA && !cellB.value && !cellC.value;
      if (isSubgroup) { try { ws.mergeCells(`A${r}:${lastColLetter}${r}`); } catch (e) { } const c = ws.getCell(`A${r}`); ws.getRow(r).eachCell(cell => { cell.font = { ...FONT, bold: true }; cell.alignment = alignC; }); c.border = { top: medium, bottom: thin, left: medium, right: medium }; if (!row.height) row.height = 14.5; continue; }
      if (isGroupTotal) { row.eachCell(cell => { cell.font = { ...FONT, bold: true }; cell.alignment = alignC; cell.border = { top: thin, bottom: medium, left: thin, right: thin }; }); continue; }
      if (isGrandTotal) { row.eachCell(cell => { cell.font = { ...FONT, bold: true }; cell.alignment = alignC; cell.border = { top: medium, bottom: medium, left: thin, right: thin }; }); if (!row.height) row.height = 14.25; continue; }
      row.eachCell((cell, col) => { const isName = col === 1; cell.font = FONT; cell.alignment = isName ? alignC : alignNum; cell.border = { top: thin, bottom: thin, left: thin, right: thin }; });
    }
    const bottomRow = ws.rowCount; const rightCol = lastColIndex + 1;
    for (let c = 1; c <= rightCol; c++) { const topCell = ws.getCell(1, c); topCell.border = { ...(topCell.border || {}), top: medium, left: c === 1 ? medium : (topCell.border?.left || thin), right: c === rightCol ? medium : (topCell.border?.right || thin) }; }
    for (let r = 1; r <= bottomRow; r++) { ws.getCell(r, 1).border = { ...(ws.getCell(r, 1).border || {}), left: medium }; ws.getCell(r, rightCol).border = { ...(ws.getCell(r, rightCol).border || {}), right: medium }; }
    for (let c = 1; c <= rightCol; c++) { const bot = ws.getCell(bottomRow, c); bot.border = { ...(bot.border || {}), bottom: medium }; }
    ws.views = [{ state: 'frozen', ySplit: headerRowIndex || 3 }];
    const buffer = await wb.xlsx.writeBuffer(); const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); showToast(`Експортовано ${filename}`, 'success');
  } catch (err) { console.error(err); showToast('Помилка ExcelJS експорту', 'error'); }
}

function applyExactVisualTemplateSheetJS(ws) {
  if (!ws) return;
  const STATIC_MERGES = ['A1:I1', 'B2:I2'];
  ws['!merges'] = ws['!merges'] || [];
  const rangeRef = ws['!ref'];
  function inRange(r) { if (!rangeRef) return true; const [s, e] = rangeRef.split(':'); const decode = XLSX.utils.decode_cell; const rs = decode(s), re = decode(e); const mr = XLSX.utils.decode_range(r); return mr.s.r >= rs.r && mr.e.r <= re.r; }
  STATIC_MERGES.forEach(m => { if (inRange(m) && !ws['!merges'].some(ex => JSON.stringify(ex) === JSON.stringify(XLSX.utils.decode_range(m)))) { try { ws['!merges'].push(XLSX.utils.decode_range(m)); } catch (e) { } } });
  const rows = []; const totalRows = (XLSX.utils.decode_range(ws['!ref']).e.r) + 1;
  for (let r = 0; r < totalRows; r++) { rows[r] = {}; }
  const setH = (idx, h) => { if (rows[idx - 1]) rows[idx - 1].hpt = h; };
  setH(1, 73.5); setH(2, 36.75); setH(3, 30); setH(4, 27);
  for (let r = 5; r <= totalRows; r++) { if (!rows[r - 1].hpt) rows[r - 1].hpt = 12.75; }
  ws['!rows'] = rows;
  const thin = { style: 'thin', color: { rgb: '000000' } }; const medium = { style: 'medium', color: { rgb: '000000' } };
  const rng = XLSX.utils.decode_range(ws['!ref']);
  function setStyle(ref, style) { if (!ws[ref]) return; ws[ref].s = Object.assign({}, ws[ref].s || {}, style); }
  setStyle('A1', { font: { name: 'Calibri', sz: 11, bold: true }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: { top: medium, bottom: medium, left: medium, right: medium } });
  const headerRow = 3; const cols = FIXED_LAYOUT.length; for (let c = 0; c < cols; c++) { const ref = XLSX.utils.encode_cell({ r: headerRow - 1, c }); setStyle(ref, { font: { name: 'Calibri', sz: 11, bold: true }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, fill: { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } }, border: { top: medium, bottom: medium, left: thin, right: thin } }); }
  for (let r = headerRow + 1; r <= rng.e.r + 1; r++) {
    const aRef = XLSX.utils.encode_cell({ r: r - 1, c: 0 }); const val = ws[aRef] ? ws[aRef].v : undefined;
    const totalColIndex = FIXED_LAYOUT.findIndex(c => c.key === 'total');
    const isTotal = val && /^всього/i.test(String(val).toLowerCase());
    const bRef = XLSX.utils.encode_cell({ r: r - 1, c: 1 }); const cRef = XLSX.utils.encode_cell({ r: r - 1, c: 2 });
    const isSub = (!ws[bRef] || ws[bRef].v === '') && (!ws[cRef] || ws[cRef].v === '') && val;
    for (let c = 0; c < cols; c++) {
      const ref = XLSX.utils.encode_cell({ r: r - 1, c }); if (!ws[ref]) continue;
      const base = { font: { name: 'Calibri', sz: 11 }, alignment: { vertical: 'center', horizontal: c === 0 ? 'left' : 'center' }, border: { top: thin, bottom: thin, left: thin, right: thin } };
      if (isSub) { base.font.bold = true; base.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }; base.fill = { patternType: 'solid', fgColor: { rgb: 'FFF2C4' } }; base.border = { top: medium, bottom: medium, left: medium, right: medium }; }
      if (isTotal) { base.font.bold = true; base.alignment = { horizontal: 'center', vertical: 'center' }; base.fill = { patternType: 'solid', fgColor: { rgb: 'D9EDF7' } }; base.border = { top: medium, bottom: medium, left: thin, right: thin }; }
      setStyle(ref, base);
    }
  }
}
