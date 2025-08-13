import { $, escapeHtml, diffMinutes, rateDisplay, parseHoursInterval } from './utils.js';
import { state, persist } from './state.js';
import { computePays, isDayOff } from './pay.js';
import { applyHoursToEmployees, parseEmployeesBulk, parseHoursBulk } from './parse.js';
import { exportExcel } from './xlsx.js';

export function showToast(msg, type = 'info', timeout = 4000) {
  let cont = $('.toast-container');
  if (!cont) {
    cont = document.createElement('div');
    cont.className = 'toast-container';
    document.body.appendChild(cont);
  }
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<div>${msg}</div><button aria-label="Закрити">✕</button>`;
  el.querySelector('button').onclick = () => {
    el.remove();
  };
  cont.appendChild(el);
  if (timeout)
    setTimeout(() => {
      el.remove();
    }, timeout);
}

export function renderEmployeesTable() {
  const container = document.getElementById('employeesTableContainer');
  if (!state.employees.length) {
    container.innerHTML = `<div class="empty">Немає співробітників. Додайте список.</div>`;
    return;
  }
  const grouped = [];
  const orderMap = new Map();
  state.employees.forEach((emp) => {
    if (isDayOff(emp)) return;
    const g = emp.group || '_NO_GROUP_';
    if (!orderMap.has(g)) orderMap.set(g, []);
    orderMap.get(g).push(emp);
  });
  const seen = new Set();
  state.employees.forEach((emp) => {
    if (isDayOff(emp)) return;
    const g = emp.group || '_NO_GROUP_';
    if (!seen.has(g) && orderMap.has(g)) {
      grouped.push(g);
      seen.add(g);
    }
  });
  if (!grouped.length) {
    container.innerHTML = `<div class="empty">Немає записів (усі вихідні)</div>`;
    return;
  }
  const rows = [];
  rows.push(
    `<table class='payroll-table'><thead><tr><th>#</th><th>ПІБ</th><th>Посада</th><th>Ставка</th><th>Години</th><th>Продажі</th><th>Подарунки</th><th>Утримано</th><th>ЗП</th></tr></thead><tbody>`
  );
  for (const g of grouped) {
    if (g !== '_NO_GROUP_')
      rows.push(`<tr class='group-row'><td colspan='9'>${escapeHtml(g)}</td></tr>`);
    for (const emp of orderMap.get(g)) {
      const fixed = emp.rateType === 'fixed';
      const rateDisp = rateDisplay(emp);
      const fixedTag = fixed ? '<span class="tag-fixed">FIX</span>' : '';
      const payInt = typeof emp.pay === 'number' ? Math.round(emp.pay) : '';
      const rawHours = emp.hoursText || '';
      const hoursEsc = escapeHtml(rawHours);
      // Pre-render validation for hours so invalid stays highlighted across renders
      let invalidHours = false;
      const clean = rawHours.replace(/\s+/g, '');
      if (clean) {
        const parsed = parseHoursInterval(rawHours);
        invalidHours = !parsed.valid;
      }
      const hoursTdClass = `editable${invalidHours ? ' invalid' : ''}`;
      rows.push(`<tr data-id='${emp.id}' class='${fixed ? 'mark-fixed' : ''}'>
        <td>${emp.order}</td>
        <td>${escapeHtml(emp.name)}</td>
        <td>${escapeHtml(emp.position)} ${fixedTag}</td>
        <td>${rateDisp}</td>
        <td class='${hoursTdClass}' data-field='hoursText' contenteditable='true' title='Формат 10:00-21:30'>${hoursEsc}</td>
        <td class='editable' data-field='sales' contenteditable='true'>${escapeHtml(String(emp.sales || ''))}</td>
        <td class='editable' data-field='gifts' contenteditable='true'>${escapeHtml(String(emp.gifts || ''))}</td>
        <td class='editable' data-field='withheld' contenteditable='true'>${escapeHtml(String(emp.withheld || ''))}</td>
        <td class='pay'>${payInt}</td>
      </tr>`);
    }
  }
  rows.push('</tbody></table>');
  container.innerHTML = rows.join('');
  // Internal scroll removed; no shadow toggle needed
  container.querySelectorAll('td.editable').forEach((cell) => {
    cell.addEventListener('blur', onCellEdit);
    cell.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        cell.blur();
      }
    });
  });
  if (!container.dataset.ctxBound) {
    container.addEventListener('contextmenu', (e) => {
      const tr = e.target.closest('tr[data-id]');
      if (!tr) return;
      e.preventDefault();
      openEmployeeContextMenu(e, tr.dataset.id);
    });
    container.dataset.ctxBound = '1';
  }
  updateExportButtonState();
}

function updateExportButtonState() {
  const btn = document.getElementById('exportExcelBtn');
  if (!btn) return;
  const anyInvalid = !!document.querySelector('#employeesTableContainer td.editable.invalid');
  const hasEmployees = state.employees.length > 0;
  if (!hasEmployees) {
    btn.disabled = true;
    btn.title = 'Немає даних для експорту';
    return;
  }
  if (anyInvalid) {
    btn.disabled = true;
    btn.title = 'Виправте некоректні години щоб експортувати';
  } else {
    btn.disabled = false;
    btn.title = 'Експортувати Excel';
  }
}

function onCellEdit(e) {
  const td = e.currentTarget;
  const tr = td.closest('tr');
  if (!tr) return;
  const id = tr.dataset.id;
  const emp = state.employees.find((x) => x.id === id);
  if (!emp) return;
  const field = td.dataset.field;
  const val = td.textContent.trim();
  if (field === 'sales' || field === 'gifts' || field === 'withheld') {
    emp[field] = parseFloat(val.replace(/,/g, '.')) || 0;
    recalcPersistRender();
    return;
  }
  if (field === 'hoursText') {
    emp.hoursText = val;
    const parsed = parseHoursInterval(val);
    if (!parsed.valid) {
      if (!td.classList.contains('invalid')) {
        td.classList.add('invalid');
        showToast('Некоректний формат часу. Використовуйте HH:MM-HH:MM', 'error', 5000);
      }
      updateExportButtonState();
      return;
    } else {
      td.classList.remove('invalid');
      emp.hoursMinutes = parsed.minutes;
      recalcPersistRender();
      updateExportButtonState();
    }
  }
}

function recalcPersistRender(toast, type = 'success') {
  computePays();
  persist();
  renderEmployeesTable();
  if (toast) showToast(toast, type);
}

export function openModal({ title, body, actions = [] }) {
  let layer = document.getElementById('modalLayer');
  if (!layer) {
    layer = document.createElement('div');
    layer.id = 'modalLayer';
    layer.className = 'modal-layer';
    document.body.appendChild(layer);
  }
  layer.innerHTML = '';
  layer.classList.remove('hidden');
  const modal = document.createElement('div');
  modal.className = 'modal';
  modal.innerHTML = `<button class='close-btn' aria-label='Закрити'>✕</button><h3>${title}</h3><div class='modal-body'></div><footer></footer>`;
  modal.querySelector('.modal-body').appendChild(body);
  const footer = modal.querySelector('footer');
  actions.forEach((a) => {
    const btn = document.createElement('button');
    btn.textContent = a.label;
    btn.className = a.class || '';
    btn.onclick = () => a.onClick && a.onClick();
    footer.appendChild(btn);
  });
  modal.querySelector('.close-btn').onclick = closeModal;
  layer.appendChild(modal);
  return modal;
}
export function closeModal() {
  const layer = document.getElementById('modalLayer');
  if (layer) layer.classList.add('hidden');
}

export function importEmployeesFlow() {
  const ta = document.createElement('textarea');
  ta.placeholder = 'Вставте список співробітників (з групами)';
  const helper = document.createElement('div');
  helper.className = 'helper';
  helper.textContent =
    'Формат: Імʼя Прізвище\tПосада\t120 грн/год. Працює також: кілька пробілів, кома, ;, |, тире між полями. Рядок без розділювачів = назва групи. Ставка: число/грн/год, "фікс 1000" або 5%.';
  const body = document.createElement('div');
  body.append(ta, helper);
  openModal({
    title: 'Імпорт співробітників',
    body,
    actions: [
      { label: 'Скасувати', class: 'subtle', onClick: closeModal },
      {
        label: 'Імпортувати',
        class: 'primary',
        onClick: () => {
          const emps = parseEmployeesBulk(ta.value.trim());
          if (!emps.length) {
            showToast('Порожній список', 'error');
            return;
          }
          state.employees = emps;
          recalcPersistRender('Імпортовано ' + emps.length);
          closeModal();
        },
      },
    ],
  });
}

export function editEmployeesFlow() {
  if (!state.employees.length) {
    showToast('Немає співробітників', 'error');
    return;
  }
  const ta = document.createElement('textarea');
  const lines = [];
  let prev = '__INIT__';
  state.employees.forEach((e) => {
    if (e.group && e.group !== prev) {
      lines.push(e.group);
      prev = e.group;
    }
    lines.push(`${e.name}\t${e.position}\t${e.rawRateStr}`);
  });
  ta.value = lines.join('\n');
  const helper = document.createElement('div');
  helper.className = 'helper';
  helper.textContent =
    'Редагуйте список. Можна використовувати таби, кілька пробілів, коми, ;, | або тире між полями. Рядок без розділювачів = група. Години/продажі збережемо за порядком.';
  const body = document.createElement('div');
  body.append(ta, helper);
  openModal({
    title: 'Редагувати список',
    body,
    actions: [
      { label: 'Відміна', class: 'subtle', onClick: closeModal },
      {
        label: 'Зберегти',
        class: 'primary',
        onClick: () => {
          const list = parseEmployeesBulk(ta.value.trim());
          if (!list.length) {
            showToast('Порожньо', 'error');
            return;
          }
          list.forEach((n) => {
            const old = state.employees.find((o) => o.order === n.order);
            if (old) {
              n.hoursText = old.hoursText;
              n.hoursMinutes = old.hoursMinutes;
              n.sales = old.sales;
              n.gifts = old.gifts;
            }
          });
          state.employees = list;
          recalcPersistRender('Оновлено');
          closeModal();
        },
      },
    ],
  });
}

export function importHoursFlow() {
  if (!state.employees.length) {
    showToast('Немає співробітників', 'error');
    return;
  }
  const ta = document.createElement('textarea');
  ta.placeholder = '10:00-21:00 кожен рядок';
  const helper = document.createElement('div');
  helper.className = 'helper';
  helper.textContent =
    'Порожні рядки ігноруються. "в" = пропуск. "вихідна" / "вибув" = службова позначка.';
  const body = document.createElement('div');
  body.append(ta, helper);
  openModal({
    title: 'Імпорт годин',
    body,
    actions: [
      { label: 'Скасувати', class: 'subtle', onClick: closeModal },
      {
        label: 'Застосувати',
        class: 'primary',
        onClick: () => {
          const lines = parseHoursBulk(ta.value.trim());
          applyHoursToEmployees(lines);
          recalcPersistRender('Години додано');
          closeModal();
          promptSalesIfNeeded();
        },
      },
    ],
  });
}

function promptSalesIfNeeded() {
  const targets = state.employees.filter(
    (e) => (e.rateType === 'waiter' || e.rateType === 'hostess') && e.hoursMinutes > 0
  );
  if (!targets.length) return;
  const body = document.createElement('div');
  body.innerHTML = '<div class="helper">Введіть продажі / подарунки</div>';
  const list = document.createElement('div');
  list.style.display = 'flex';
  list.style.flexDirection = 'column';
  list.style.gap = '10px';
  targets.forEach((emp) => {
    const row = document.createElement('div');
    row.className = 'inline-fields';
    row.innerHTML = `<div class='input'><span>${escapeHtml(emp.name)} Продажі</span><input type='number' value='${emp.sales || ''}' data-id='${emp.id}' data-field='sales'></div><div class='input'><span>Подарунки</span><input type='number' value='${emp.gifts || ''}' data-id='${emp.id}' data-field='gifts'></div>`;
    list.appendChild(row);
  });
  body.appendChild(list);
  openModal({
    title: 'Продажі',
    body,
    actions: [
      {
        label: 'OK',
        class: 'primary',
        onClick: () => {
          list.querySelectorAll('input').forEach((inp) => {
            const emp = state.employees.find((e) => e.id === inp.dataset.id);
            if (emp) emp[inp.dataset.field] = parseFloat(inp.value) || 0;
          });
          recalcPersistRender('Збережено');
          closeModal();
        },
      },
    ],
  });
}

export function clearHours() {
  state.employees.forEach((e) => {
    e.hoursText = '';
    e.hoursMinutes = 0;
    e.sales = 0;
    e.gifts = 0;
    e.withheld = 0;
    e.pay = 0;
  });
  recalcPersistRender('Очищено години / продажі / утримання');
}
export function fullReset() {
  if (!confirm('Очистити всі дані?')) return;
  localStorage.removeItem('payroll_employees_v1');
  localStorage.removeItem('payroll_meta_v1');
  state.employees = [];
  renderEmployeesTable();
  showToast('Скидання виконано', 'success');
}
export function toggleTheme() {
  state.meta.theme = state.meta.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.classList.toggle('light', state.meta.theme === 'light');
  persist();
}

export function bindGlobalEvents() {
  $('#importEmployeesBtn').onclick = importEmployeesFlow;
  $('#editEmployeesBtn').onclick = editEmployeesFlow;
  $('#inputHoursBtn').onclick = importHoursFlow;
  $('#exportExcelBtn').onclick = exportExcel;
  $('#clearHoursBtn').onclick = clearHours;
  $('#fullResetBtn').onclick = fullReset;
  $('#themeToggleBtn').onclick = toggleTheme;
  // Initial state of export button (in case of persisted invalid data)
  requestAnimationFrame(updateExportButtonState);
  const cityInput = document.getElementById('cityInput');
  const dateInput = document.getElementById('reportDateInput');
  if (cityInput) {
    cityInput.value = state.settings.city;
    cityInput.addEventListener('change', () => {
      state.settings.city = cityInput.value.trim();
      persist();
    });
  }
  if (dateInput) {
    // Always show yesterday by default
    const d = new Date();
    d.setDate(d.getDate() - 1);
    const yest = d.toISOString().slice(0, 10);
    state.settings.reportDate = yest;
    dateInput.value = yest;
    dateInput.addEventListener('change', () => {
      state.settings.reportDate = dateInput.value;
      // Recompute because fixed per-day depends on days in month
      recalcPersistRender();
    });
  }
}

export function openEmployeeContextMenu(e, id) {
  const emp = state.employees.find((x) => x.id === id);
  if (!emp) return;
  document.querySelectorAll('.context-menu').forEach((m) => m.remove());
  const menu = document.createElement('div');
  menu.className = 'context-menu';
  function add(label, cb) {
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.onclick = () => {
      cb();
      menu.remove();
      document.removeEventListener('click', outside);
    };
    menu.appendChild(b);
  }
  add('Погодинна ставка', () => {
    const v = prompt('Ставка /год?', emp.hourlyRate || '');
    if (v !== null) {
      const r = parseFloat(v.replace(/,/g, '.'));
      if (!isNaN(r)) {
        emp.rateType = 'hourly';
        emp.hourlyRate = r;
        emp.basePay = 0;
        recalcPersistRender();
      }
    }
  });
  add('Офіціант (5%)', () => {
    emp.rateType = 'waiter';
    emp.waiterPercent = 5;
    emp.hourlyRate = 0;
    emp.basePay = 0;
    recalcPersistRender();
  });
  add('Хостес (+2%)', () => {
    const v = prompt('Погодинна ставка хостес?', emp.hourlyRate || '');
    if (v !== null) {
      const r = parseFloat(v.replace(/,/g, '.'));
      if (!isNaN(r)) {
        emp.rateType = 'hostess';
        emp.hourlyRate = r;
        emp.hostessPercent = 2;
        emp.basePay = 0;
        recalcPersistRender();
      }
    }
  });
  add('Фіксована сума', () => {
    const v = prompt('Зарплата за місяць (фікс)?', emp.monthlyBase || emp.basePay || '');
    if (v !== null) {
      const r = parseFloat(v.replace(/,/g, '.'));
      if (!isNaN(r)) {
        emp.rateType = 'fixed';
        emp.monthlyBase = r;
        // Keep legacy basePay for compatibility; per-day will be derived from monthlyBase
        emp.basePay = emp.basePay || 0;
        emp.hourlyRate = 0;
        recalcPersistRender();
      }
    }
  });
  if (emp.rateType === 'fixed')
    add('Прибрати FIX', () => {
      emp.rateType = 'hourly';
      emp.hourlyRate = emp.basePay || emp.hourlyRate;
      emp.basePay = 0;
      emp.monthlyBase = 0;
      recalcPersistRender();
    });
  let x = e.pageX + 6;
  let y = e.pageY + 6;
  menu.style.position = 'absolute';
  menu.style.visibility = 'hidden';
  menu.style.top = y + 'px';
  menu.style.left = x + 'px';
  document.body.appendChild(menu);
  const rect = menu.getBoundingClientRect();
  const vpW = window.innerWidth;
  const vpH = window.innerHeight;
  if (rect.right > vpW) {
    x = Math.max(4, x - (rect.right - vpW));
  }
  if (rect.bottom > vpH) {
    y = Math.max(4, y - (rect.bottom - vpH));
  }
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.style.visibility = 'visible';
  function outside(ev) {
    if (!menu.contains(ev.target)) {
      menu.remove();
      document.removeEventListener('click', outside);
    }
  }
  setTimeout(() => document.addEventListener('click', outside), 0);
}
