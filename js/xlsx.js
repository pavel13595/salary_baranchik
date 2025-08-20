import { state } from './state.js';
import { showToast } from './ui.js';
import { computePays } from './pay.js';
import { rateDisplay, parseHoursInterval, fixedPerDay } from './utils.js';

const FIXED_LAYOUT = [
  { key: 'name', title: '', width: 28 },
  { key: 'position', title: 'Посада', width: 16 },
  // Fixed header wording: remove duplicated word and trailing artifacts
  { key: 'hours', title: 'Кількість відпрацьованих годин', width: 18 },
  { key: 'rate', title: 'Ставка', width: 10 },
  { key: 'sales', title: 'Продажі', width: 12 },
  { key: 'withheld', title: 'Утримано', width: 12 },
  { key: 'issued', title: 'Видано', width: 12 },
  { key: 'total', title: 'Всього нараховано', width: 18 },
  { key: 'sign', title: 'Підпис отримувача', width: 18 },
];

const SUBGROUPS = [
  { name: 'Адмін. Персонал', patterns: [/^керуюч(ий|а)$/i, /^менеджер$/i, /^шеф$/i] },
  { name: 'Бар', patterns: [/бармен/i, /бармени/i] },
  { name: 'Кухня', patterns: [/кухар/i, /су[- ]?шеф/i, /суш[- ]?шеф/i] },
  { name: 'Офіціанти / ранери', patterns: [/офіціант/i, /официант/i, /ранер/i] },
  {
    name: 'Хостес / Доставка',
    patterns: [/хостес/i, /пакувальниц/i, /пакувальник/i, /упаковщ/i, /упаковщик/i],
  },
  { name: 'Господарка', patterns: [/господарк/i, /господин/i] },
  { name: 'Інший персонал', patterns: [/кур[’'`]?єр/i, /курьер/i, /завгосп/i] },
];
function classifySubgroup(position) {
  const pos = (position || '').toLowerCase();
  for (const sg of SUBGROUPS) {
    if (sg.patterns.some((rx) => rx.test(pos))) return sg.name;
  }
  return 'Інший персонал';
}

function buildReportData() {
  const repDateStr =
    state.settings.reportDate ||
    (() => {
      const d = new Date();
      d.setDate(d.getDate() - 1);
      return d.toISOString().slice(0, 10);
    })();
  const dParts = repDateStr.split('-');
  const dateStrDisp = `${dParts[2]}.${dParts[1]}`;
  const titleLine = `Той Самий Баранчик${state.settings.city ? ' ' + state.settings.city : ''}`;
  const headersLocal = FIXED_LAYOUT.map((c) => c.title);
  const data = [];
  const merges = [];
  const formulaRows = []; // { rowNumber, formula, percentRate }
  data.push([titleLine, ...Array(headersLocal.length - 1).fill('')]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: headersLocal.length - 1 } });
  const dateRow = [dateStrDisp, ...Array(headersLocal.length - 1).fill('')];
  data.push(dateRow);
  merges.push({ s: { r: 1, c: 1 }, e: { r: 1, c: headersLocal.length - 1 } });
  data.push(headersLocal);
  // Ensure header titles explicitly for columns (some earlier transformations may override)
  const headerIndex = data.length - 1;
  if (data[headerIndex]) {
    data[headerIndex][4] = 'Продажі';
    data[headerIndex][5] = 'Утримано';
    data[headerIndex][6] = 'Видано';
  }
  let rowIdx = 3;
  const isDayOff = (emp) => /вихід|вибув/i.test(emp.hoursText || '');
  const totalColIndex = FIXED_LAYOUT.findIndex((c) => c.key === 'total');
  const subgroupMap = new Map();
  SUBGROUPS.forEach((sg) => subgroupMap.set(sg.name, []));
  for (const emp of state.employees) {
    if (isDayOff(emp)) continue;
    if (!(emp.rateType === 'fixed' || emp.hoursMinutes > 0)) continue;
    const sg = classifySubgroup(emp.position);
    if (!subgroupMap.has(sg)) subgroupMap.set(sg, []);
    subgroupMap.get(sg).push(emp);
  }
  function rateDisp(e) {
    // Return numeric rate (hourly/fixed). For waiter percent we store decimal (0.05) and later format as percent.
    if (e.rateType === 'waiter') return Number(e.waiterPercent || 5) / 100;
    if (e.rateType === 'fixed') return Number(fixedPerDay(e) || 0);
    return Number(e.hourlyRate || 0);
  }
  function hoursF(e) {
    if (e.hoursMinutes || e.hoursMinutes === 0) return +(e.hoursMinutes / 60).toFixed(2);
    return 0; // ensure numeric for formulas
  }
  let subtotalFormulaTargets = []; // store subtotal row numbers for grand total
  for (const sg of SUBGROUPS.map((s) => s.name)) {
    const emps = subgroupMap.get(sg) || [];
    if (!emps.length) continue;
    data.push([sg, ...Array(headersLocal.length - 1).fill('')]);
    merges.push({ s: { r: rowIdx, c: 0 }, e: { r: rowIdx, c: headersLocal.length - 1 } });
    rowIdx++;
    const employeeRowNumbers = [];
    for (const e of emps) {
      const row = [];
      row.push(e.name); // A
      row.push(e.position); // B
      // Fixed rate employees always counted as 1 hour per requirement
      const hoursVal = e.rateType === 'fixed' ? 1 : hoursF(e);
      row.push(hoursVal); // C hours numeric (fixed -> 1)
      const rVal = rateDisp(e);
      row.push(rVal); // D rate (hourly/fixed or percent decimal for waiter)
      // Sales (net): only for roles with sales impact; subtract gifts (подарки) if such fields exist
      let rawSales = 0;
      if (
        e.rateType === 'waiter' ||
        /бар/i.test(e.position) ||
        /бармен/i.test(e.position) ||
        /хостес/i.test(e.position)
      ) {
        rawSales = Number(e.sales || 0);
      }
      const gifts = Number(
        e.gifts ??
          e.gift ??
          e.giftAmount ??
          e.presents ??
          e.presentAmount ??
          e.podarok ??
          e.podarki ??
          0
      );
  const netSales = rawSales - (isNaN(gifts) ? 0 : gifts);
  row.push(netSales > 0 ? netSales : 0); // E net sales (never negative)
      const withheldVal = Number(e.withheld || 0);
      row.push(withheldVal); // F withheld numeric (0 allowed)
      row.push(0); // G issued numeric default 0
      // Placeholder for total (H) will be formula later
      row.push('');
      row.push(''); // I sign
      data.push(row);
      const excelRowNumber = rowIdx + 1; // data index to Excel row: index 0 => row1
      // Determine subgroup specific formula
      const subgroupName = sg;
      const pos = (e.position || '').toLowerCase();
      function f(col) {
        return col + excelRowNumber;
      }
      let formula;
      const isRunner = /ранер/.test(pos);
      const isWaiter = /офіціант|официант/.test(pos);
      const isHostess = /хостес/.test(pos);
      const isPacker = /пакувальниц|пакувальник|упаковщ|упаковщик/.test(pos);
      const isCourier = /кур[’'`]?єр|курьер/.test(pos);
      // Base patterns using columns: C hours, D rate, E sales, F withheld, G issued
      if (subgroupName === 'Бар') {
        formula = `(${f('D')}*${f('C')})-${f('F')}-${f('G')}+${f('E')}*0.05`;
      } else if (subgroupName === 'Адмін. Персонал') {
        formula = `${f('D')}*${f('C')}-${f('F')}-${f('G')}`;
      } else if (subgroupName === 'Кухня') {
        formula = `${f('D')}*${f('C')}-${f('F')}-${f('G')}`;
      } else if (subgroupName === 'Офіціанти / ранери') {
        if (isWaiter) {
          // Always embed guarantee logic if enabled: IF(netSales < 10000, 500, netSales * percent) - issued - withheld
          if (e.waiterMinGuarantee !== false) {
            formula = `IF(${f('E')}<10000,500,${f('E')}*${f('D')})-${f('G')}-${f('F')}`;
          } else {
            formula = `${f('E')}*${f('D')}-${f('G')}-${f('F')}`;
          }
        } else if (isRunner) {
          formula = `${f('C')}*${f('D')}-${f('G')}-${f('F')}`;
        } else {
          formula = `${f('C')}*${f('D')}-${f('F')}-${f('G')}`;
        }
      } else if (subgroupName === 'Хостес / Доставка') {
        if (isHostess) {
          formula = `${f('D')}*${f('C')}+${f('E')}*0.02-${f('F')}-${f('G')}`;
        } else if (isPacker) {
          formula = `${f('C')}*${f('D')}-${f('G')}-${f('F')}`;
        } else {
          formula = `${f('C')}*${f('D')}-${f('F')}-${f('G')}`;
        }
      } else if (subgroupName === 'Господарка') {
        formula = `${f('C')}*${f('D')}-${f('G')}-${f('F')}`;
      } else if (subgroupName === 'Інший персонал') {
        if (isCourier) {
          formula = `${f('C')}*${f('D')}-${f('G')}-${f('F')}`;
        } else {
          formula = `${f('D')}*${f('C')}-${f('F')}-${f('G')}`;
        }
      } else {
        formula = `${f('D')}*${f('C')}-${f('F')}-${f('G')}`;
      }
      formulaRows.push({
        rowNumber: excelRowNumber,
        formula,
        percentRate: e.rateType === 'waiter',
      });
      employeeRowNumbers.push(excelRowNumber);
      rowIdx++;
    }
    if (employeeRowNumbers.length) {
      const subtotalLabel = `всього ${sg.toLowerCase()}`;
      const subtotalRow = Array(headersLocal.length).fill('');
      subtotalRow[0] = subtotalLabel;
      // placeholder for formula in H
      data.push(subtotalRow);
      const subtotalExcelRow = rowIdx + 1;
      const firstEmp = Math.min(...employeeRowNumbers);
      const lastEmp = Math.max(...employeeRowNumbers);
      const subtotalFormula = `SUM(H${firstEmp}:H${lastEmp})`;
      formulaRows.push({ rowNumber: subtotalExcelRow, formula: subtotalFormula });
      subtotalFormulaTargets.push(subtotalExcelRow);
      const issuedIdx = FIXED_LAYOUT.findIndex((c) => c.key === 'issued');
      merges.push({ s: { r: rowIdx, c: 1 }, e: { r: rowIdx, c: issuedIdx } });
      rowIdx++;
    }
  }
  const grand = Array(headersLocal.length).fill('');
  grand[0] = 'ВСЬОГО';
  data.push(grand);
  const grandExcelRow = rowIdx + 1;
  if (subtotalFormulaTargets.length) {
    const parts = subtotalFormulaTargets.map((rn) => `H${rn}`);
    const grandFormula = `SUM(${parts.join(',')})`;
    formulaRows.push({ rowNumber: grandExcelRow, formula: grandFormula });
  }
  const issuedIdx = FIXED_LAYOUT.findIndex((c) => c.key === 'issued');
  merges.push({ s: { r: rowIdx, c: 1 }, e: { r: rowIdx, c: issuedIdx } });
  return { headers: headersLocal, data, merges, repDateStr, formulaRows };
}

// (Removed comma enforcement: keeping numeric values; Excel will display per user locale. For explicit formatting we set numFmt later.)

export async function exportExcel() {
  // Safety: ensure latest calculations (in case something changed without re-render)
  computePays();
  // Block export if any invalid hours exist (data integrity safeguard)
  try {
    let hoursInvalid = state.employees.some((emp) => !parseHoursInterval(emp.hoursText).valid);
    // Fallback: if UI shows no invalid cells, allow export (prevents false positives)
    if (hoursInvalid) {
      const anyDomInvalid =
        typeof document !== 'undefined' &&
        document.querySelector('#employeesTableContainer td.editable.invalid');
      if (!anyDomInvalid) hoursInvalid = false;
    }
    if (hoursInvalid) {
      showToast('Є некоректно введений час — виправте перед експортом', 'error', 6000);
      return;
    }
  } catch (e) {
    /* fail-safe: if validation throws, continue with export attempt */
  }
  if (window.ExcelJS) {
    return exportExcelExcelJS();
  }
  if (!state.employees.length) {
    showToast('Немає даних', 'error');
    return;
  }
  if (!window.XLSX) {
    showToast('XLSX не завантажено', 'error');
    return;
  }
  const { headers, data, merges, repDateStr, formulaRows } = buildReportData();
  try {
    const ws = XLSX.utils.aoa_to_sheet(data);
    if (merges.length) ws['!merges'] = merges;
    ws['!cols'] = FIXED_LAYOUT.map((c) => ({ wch: c.width }));
    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: data.length - 1, c: headers.length - 1 },
    });
    // Apply formulas (SheetJS): set total column (H) formulas
    formulaRows.forEach((fr) => {
      const cellRef = 'H' + fr.rowNumber;
      if (!ws[cellRef]) ws[cellRef] = { t: 'n', f: fr.formula };
      else ws[cellRef].f = fr.formula;
    });
    // Number formatting: hours (C) -> 0.00 ; rate (D) integer except waiter percent -> 0% ; sales (E), withheld (F), issued (G), total (H) -> integer
    const rng = XLSX.utils.decode_range(ws['!ref']);
    const waiterRateRows = new Set(
      formulaRows.filter((fr) => fr.percentRate).map((fr) => fr.rowNumber)
    );
    for (let r = 0; r <= rng.e.r; r++) {
      const excelRow = r + 1;
      const cHours = XLSX.utils.encode_cell({ r, c: 2 });
      if (ws[cHours] && typeof ws[cHours].v === 'number') ws[cHours].z = '0.00';
      const cRate = XLSX.utils.encode_cell({ r, c: 3 });
      if (ws[cRate] && typeof ws[cRate].v === 'number')
        ws[cRate].z = waiterRateRows.has(excelRow) ? '0%' : '0';
      const cSales = XLSX.utils.encode_cell({ r, c: 4 });
      if (ws[cSales] && typeof ws[cSales].v === 'number') ws[cSales].z = '0;-0;;';
      const cWithheld = XLSX.utils.encode_cell({ r, c: 5 });
      if (ws[cWithheld] && typeof ws[cWithheld].v === 'number') ws[cWithheld].z = '0;-0;;';
      const cIssued = XLSX.utils.encode_cell({ r, c: 6 });
      if (ws[cIssued] && typeof ws[cIssued].v === 'number') ws[cIssued].z = '0;-0;;';
      const cTotal = XLSX.utils.encode_cell({ r, c: 7 });
      if (
        ws[cTotal] &&
        (ws[cTotal].v === undefined || ws[cTotal].f || typeof ws[cTotal].v === 'number')
      )
        ws[cTotal].z = '0';
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Відомість');
    applyExactVisualTemplateSheetJS(ws);
    const citySegment = state.settings.city ? state.settings.city.replace(/\s+/g, '_') : 'Місто';
    const dParts = repDateStr.split('-');
    const dateSeg = `${dParts[2]}.${dParts[1]}`;
    const filename = `Відомість_${citySegment}_${dateSeg}.xlsx`;
    XLSX.writeFile(wb, filename, { compression: true });
    showToast(`Експортовано (обмежений стиль) ${filename}`, 'success');
  } catch (err) {
    console.error(err);
    showToast('Помилка експорту', 'error');
  }
}

async function exportExcelExcelJS() {
  computePays();
  // Duplicate guard (in case exportExcelExcelJS invoked directly or by future refactor)
  try {
    let hoursInvalid = state.employees.some((emp) => !parseHoursInterval(emp.hoursText).valid);
    if (hoursInvalid) {
      const anyDomInvalid =
        typeof document !== 'undefined' &&
        document.querySelector('#employeesTableContainer td.editable.invalid');
      if (!anyDomInvalid) hoursInvalid = false;
    }
    if (hoursInvalid) {
      showToast('Є некоректно введений час — виправте перед експортом', 'error', 6000);
      return;
    }
  } catch (_) {}
  if (!state.employees.length) {
    showToast('Немає даних', 'error');
    return;
  }
  const { data, merges, repDateStr, formulaRows } = buildReportData();
  const citySegment = state.settings.city ? state.settings.city.replace(/\s+/g, '_') : 'Місто';
  const dParts = repDateStr.split('-');
  const dateSeg = `${dParts[2]}.${dParts[1]}`;
  const filename = `Відомість_${citySegment}_${dateSeg}.xlsx`;
  try {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Відомість', { properties: { defaultRowHeight: 12.75 } });
    data.forEach((r) => ws.addRow(r));
    // Apply formulas and numeric formats in ExcelJS
    formulaRows.forEach((fr) => {
      const cell = ws.getCell('H' + fr.rowNumber);
      cell.value = { formula: fr.formula };
      cell.numFmt = '0';
    });
    // Column formats
    if (ws.getColumn(3)) ws.getColumn(3).numFmt = '0.00'; // hours keep 2 decimals
    if (ws.getColumn(4)) ws.getColumn(4).numFmt = '0'; // rate integer
    if (ws.getColumn(5)) ws.getColumn(5).numFmt = '0;-0;;'; // sales hide zero
    if (ws.getColumn(6)) ws.getColumn(6).numFmt = '0;-0;;'; // withheld hide zero
    if (ws.getColumn(7)) ws.getColumn(7).numFmt = '0;-0;;'; // issued hide zero
    if (ws.getColumn(8)) ws.getColumn(8).numFmt = '0'; // total
    // Percent rates for waiter rows: set individual cell format to percent (override integer)
    formulaRows
      .filter((fr) => fr.percentRate)
      .forEach((fr) => {
        const cell = ws.getCell('D' + fr.rowNumber);
        cell.numFmt = '0%';
      });
    function colLetter(idx) {
      let s = '';
      idx = idx + 1;
      while (idx > 0) {
        const m = (idx - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        idx = ((idx - 1) / 26) | 0;
      }
      return s;
    }
    const lastColIndex = FIXED_LAYOUT.length - 1;
    const lastColLetter = colLetter(lastColIndex);
    const dynMerges = merges.map(
      (m) => `${colLetter(m.s.c)}${m.s.r + 1}:${colLetter(m.e.c)}${m.e.r + 1}`
    );
    const STATIC_MERGES = ['A1:' + lastColLetter + '1', 'B2:' + lastColLetter + '2'];
    const maxRow = ws.rowCount;
    function mergeSafe(range) {
      try {
        const rowNums = range.match(/\d+/g);
        if (rowNums) {
          const maxR = Math.max(...rowNums.map(Number));
          if (maxR > maxRow) return;
        }
        ws.mergeCells(range);
      } catch (e) {}
    }
    [...dynMerges, ...STATIC_MERGES].forEach(mergeSafe);
    const setW = (col, val) => {
      if (ws.getColumn(col)) ws.getColumn(col).width = val;
    };
    setW('A', 30.7109375);
    setW('B', 19.5703125);
    setW('C', 9.28515625);
    setW('G', 11.5703125);
    setW('H', 12.140625);
    setW('I', 14.7109375);
    setW('J', 14.42578125);
    setW('K', 14.42578125);
    ['D', 'E', 'F'].forEach((c) => {
      if (ws.getColumn(c)) ws.getColumn(c).width = 9.3;
    });
    const keyHeights = { 1: 12.75, 2: 12.75, 3: 65, 4: 12.75 };
    Object.entries(keyHeights).forEach(([r, h]) => {
      if (ws.getRow(+r)) ws.getRow(+r).height = h;
    });
    for (let r = 5; r <= ws.rowCount; r++) {
      if (!ws.getRow(r).height) ws.getRow(r).height = 12.75;
    }
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
        if (v === 'Посада') {
          headerRowIndex = r;
          r = ws.rowCount + 1;
          break;
        }
      }
    }
    ws.getCell('A1').font = { ...FONT, bold: true };
    ws.getCell('A1').alignment = alignC;
    ws.getRow(1).eachCell((cell) => {
      cell.border = { ...(cell.border || {}), bottom: medium, top: cell.border?.top };
    });
    if (ws.getRow(2)) {
      ws.getRow(2).eachCell((c) => {
        c.font = { ...FONT, bold: true };
        c.alignment = alignC;
      });
    }
    if (headerRowIndex) {
      const hr = ws.getRow(headerRowIndex);
      hr.eachCell((cell) => {
        cell.font = { ...FONT, bold: true };
        cell.alignment = alignC;
        cell.border = { top: medium, bottom: medium, left: thin, right: thin };
      });
      // Force header labels for E/F/G (sales/withheld/issued) and set text format to avoid hiding
      const labels = [
        ['E', 'Продажі'],
        ['F', 'Утримано'],
        ['G', 'Видано'],
      ];
      labels.forEach(([col, text]) => {
        const c = ws.getCell(col + headerRowIndex);
        c.value = text;
        c.numFmt = '@';
      });
    }
    for (let r = 1; r <= ws.rowCount; r++) {
      if (r === 1 || r === 2 || r === headerRowIndex) continue;
      const row = ws.getRow(r);
      const valA = String(ws.getCell(`A${r}`).value || '').trim();
      const cellB = ws.getCell(`B${r}`);
      const cellC = ws.getCell(`C${r}`);
      const isGrandTotal = /^всього$/i.test(valA) || /^ВСЬОГО$/.test(valA);
      const isGroupTotal = /^всього\s+/.test(valA.toLowerCase()) && !isGrandTotal;
      const isSubgroup = !isGrandTotal && !isGroupTotal && !!valA && !cellB.value && !cellC.value;
      if (isSubgroup) {
        try {
          ws.mergeCells(`A${r}:${lastColLetter}${r}`);
        } catch (e) {}
        const c = ws.getCell(`A${r}`);
        ws.getRow(r).eachCell((cell) => {
          cell.font = { ...FONT, bold: true };
          cell.alignment = alignC;
        });
        c.border = { top: medium, bottom: thin, left: medium, right: medium };
        if (!row.height) row.height = 14.5;
        continue;
      }
      if (isGroupTotal) {
        row.eachCell((cell) => {
          cell.font = { ...FONT, bold: true };
          cell.alignment = alignC;
          cell.border = { top: thin, bottom: medium, left: thin, right: thin };
        });
        continue;
      }
      if (isGrandTotal) {
        row.eachCell((cell) => {
          cell.font = { ...FONT, bold: true };
          cell.alignment = alignC;
          cell.border = { top: medium, bottom: medium, left: thin, right: thin };
        });
        if (!row.height) row.height = 14.25;
        continue;
      }
      row.eachCell((cell, col) => {
        const isName = col === 1;
        cell.font = FONT;
        cell.alignment = isName ? alignC : alignNum;
        cell.border = { top: thin, bottom: thin, left: thin, right: thin };
      });
    }
    const bottomRow = ws.rowCount;
    const rightCol = lastColIndex + 1;
    for (let c = 1; c <= rightCol; c++) {
      const topCell = ws.getCell(1, c);
      topCell.border = {
        ...(topCell.border || {}),
        top: medium,
        left: c === 1 ? medium : topCell.border?.left || thin,
        right: c === rightCol ? medium : topCell.border?.right || thin,
      };
    }
    for (let r = 1; r <= bottomRow; r++) {
      ws.getCell(r, 1).border = { ...(ws.getCell(r, 1).border || {}), left: medium };
      ws.getCell(r, rightCol).border = { ...(ws.getCell(r, rightCol).border || {}), right: medium };
    }
    for (let c = 1; c <= rightCol; c++) {
      const bot = ws.getCell(bottomRow, c);
      bot.border = { ...(bot.border || {}), bottom: medium };
    }
    ws.views = [{ state: 'frozen', ySplit: headerRowIndex || 3 }];
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    showToast(`Експортовано ${filename}`, 'success');
  } catch (err) {
    console.error(err);
    showToast('Помилка ExcelJS експорту', 'error');
  }
}

function applyExactVisualTemplateSheetJS(ws) {
  if (!ws) return;
  const STATIC_MERGES = ['A1:I1', 'B2:I2'];
  ws['!merges'] = ws['!merges'] || [];
  const rangeRef = ws['!ref'];
  function inRange(r) {
    if (!rangeRef) return true;
    const [s, e] = rangeRef.split(':');
    const decode = XLSX.utils.decode_cell;
    const rs = decode(s),
      re = decode(e);
    const mr = XLSX.utils.decode_range(r);
    return mr.s.r >= rs.r && mr.e.r <= re.r;
  }
  STATIC_MERGES.forEach((m) => {
    if (
      inRange(m) &&
      !ws['!merges'].some((ex) => JSON.stringify(ex) === JSON.stringify(XLSX.utils.decode_range(m)))
    ) {
      try {
        ws['!merges'].push(XLSX.utils.decode_range(m));
      } catch (e) {}
    }
  });
  const rows = [];
  const totalRows = XLSX.utils.decode_range(ws['!ref']).e.r + 1;
  for (let r = 0; r < totalRows; r++) {
    rows[r] = {};
  }
  const setH = (idx, h) => {
    if (rows[idx - 1]) rows[idx - 1].hpt = h;
  };
  setH(1, 73.5);
  setH(2, 36.75);
  setH(3, 30);
  setH(4, 27);
  for (let r = 5; r <= totalRows; r++) {
    if (!rows[r - 1].hpt) rows[r - 1].hpt = 12.75;
  }
  ws['!rows'] = rows;
  const thin = { style: 'thin', color: { rgb: '000000' } };
  const medium = { style: 'medium', color: { rgb: '000000' } };
  const rng = XLSX.utils.decode_range(ws['!ref']);
  function setStyle(ref, style) {
    if (!ws[ref]) return;
    ws[ref].s = Object.assign({}, ws[ref].s || {}, style);
  }
  setStyle('A1', {
    font: { name: 'Calibri', sz: 11, bold: true },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: { top: medium, bottom: medium, left: medium, right: medium },
  });
  const headerRow = 3;
  const cols = FIXED_LAYOUT.length;
  for (let c = 0; c < cols; c++) {
    const ref = XLSX.utils.encode_cell({ r: headerRow - 1, c });
    setStyle(ref, {
      font: { name: 'Calibri', sz: 11, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      fill: { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } },
      border: { top: medium, bottom: medium, left: thin, right: thin },
    });
  }
  for (let r = headerRow + 1; r <= rng.e.r + 1; r++) {
    const aRef = XLSX.utils.encode_cell({ r: r - 1, c: 0 });
    const val = ws[aRef] ? ws[aRef].v : undefined;
    const totalColIndex = FIXED_LAYOUT.findIndex((c) => c.key === 'total');
    const isTotal = val && /^всього/i.test(String(val).toLowerCase());
    const bRef = XLSX.utils.encode_cell({ r: r - 1, c: 1 });
    const cRef = XLSX.utils.encode_cell({ r: r - 1, c: 2 });
    const isSub = (!ws[bRef] || ws[bRef].v === '') && (!ws[cRef] || ws[cRef].v === '') && val;
    for (let c = 0; c < cols; c++) {
      const ref = XLSX.utils.encode_cell({ r: r - 1, c });
      if (!ws[ref]) continue;
      const base = {
        font: { name: 'Calibri', sz: 11 },
        alignment: { vertical: 'center', horizontal: c === 0 ? 'left' : 'center' },
        border: { top: thin, bottom: thin, left: thin, right: thin },
      };
      if (isSub) {
        base.font.bold = true;
        base.alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
        base.fill = { patternType: 'solid', fgColor: { rgb: 'FFF2C4' } };
        base.border = { top: medium, bottom: medium, left: medium, right: medium };
      }
      if (isTotal) {
        base.font.bold = true;
        base.alignment = { horizontal: 'center', vertical: 'center' };
        base.fill = { patternType: 'solid', fgColor: { rgb: 'D9EDF7' } };
        base.border = { top: medium, bottom: medium, left: thin, right: thin };
      }
      setStyle(ref, base);
    }
  }
}
