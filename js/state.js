// State management and persistence
/**
 * @typedef {import('./utils.js').Employee} Employee
 * @typedef {Object} AppState
 * @property {Employee[]} employees
 * @property {{ theme: 'dark'|'light', lastUpdated: string|null }} meta
 * @property {{ columns: any[], groupSubtotalMergeUntil: string, dateRow: boolean }} layout
 * @property {{ city: string, reportDate: string, showOfficial:boolean }} settings
 */
export const STORAGE_KEY = 'payroll_employees_v1';
export const STORAGE_META = 'payroll_meta_v1';
export const LAYOUT_KEY = 'payroll_excel_layout_v1';
export const STORAGE_SETTINGS = 'payroll_settings_v1';

/** @type {AppState} */
export const state = {
  employees: [],
  meta: { theme: 'dark', lastUpdated: null },
  layout: {
    columns: [
      { key: 'name', title: 'ПІБ', width: 28, enabled: true },
      { key: 'position', title: 'Посада', width: 16, enabled: true },
      { key: 'hours', title: 'Кількість відпрацьованих годин', width: 18, enabled: true },
      { key: 'rate', title: 'Ставка', width: 10, enabled: true },
      { key: 'sales', title: 'Продажі', width: 12, enabled: true },
      { key: 'gifts', title: 'Утримано', width: 12, enabled: true },
      { key: 'issued', title: 'Видано', width: 12, enabled: true },
      { key: 'total', title: 'Всього нараховано', width: 18, enabled: true },
      { key: 'sign', title: 'Підпис отримувача', width: 18, enabled: true },
    ],
    groupSubtotalMergeUntil: 'issued',
    dateRow: true,
  },
  settings: { city: '', reportDate: '', showOfficial: false },
};

export const defaultLayout = JSON.parse(JSON.stringify(state.layout));

export function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) state.employees = JSON.parse(saved);
    const meta = localStorage.getItem(STORAGE_META);
    if (meta) state.meta = { ...state.meta, ...JSON.parse(meta) };
    const layout = localStorage.getItem(LAYOUT_KEY);
    if (layout) state.layout = { ...state.layout, ...JSON.parse(layout) };
    const settings = localStorage.getItem(STORAGE_SETTINGS);
    if (settings) state.settings = { ...state.settings, ...JSON.parse(settings) };
    if (typeof state.settings.showOfficial !== 'boolean') state.settings.showOfficial = false;
  } catch (e) {
    console.error('Помилка завантаження стану', e);
  }
  // Always default to yesterday's date on each load
  {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    state.settings.reportDate = d.toISOString().slice(0, 10);
  }
  document.documentElement.classList.toggle('light', state.meta.theme === 'light');
}

export function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.employees));
  state.meta.lastUpdated = new Date().toISOString();
  localStorage.setItem(STORAGE_META, JSON.stringify(state.meta));
  localStorage.setItem(LAYOUT_KEY, JSON.stringify(state.layout));
  localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(state.settings));
}
