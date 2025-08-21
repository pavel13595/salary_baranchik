// Entry point - wires modules
import { loadState, persist, state } from './state.js';
import { computePays } from './pay.js';
import { renderEmployeesTable, bindGlobalEvents, showToast } from './ui.js';

// Build/version marker (update hash when deploying to GitHub Pages to force cache refresh)
const APP_VERSION = '2025.08.21.3';
console.info('[Payroll] App version', APP_VERSION);

loadState();
computePays();
renderEmployeesTable();
bindGlobalEvents();

// === Live update checker (feature #20) ===
const VERSION_CHECK_INTERVAL_MS = 300_000; // 5 minutes
async function checkForUpdate() {
  try {
    const res = await fetch('js/app.js?cv=' + Date.now(), { cache: 'no-store' });
    if (!res.ok) return;
    const txt = await res.text();
    const m = txt.match(/APP_VERSION\s*=\s*'([^']+)'/);
    if (m && m[1] && m[1] !== APP_VERSION) {
      if (!document.getElementById('updatePrompt')) {
        const bar = document.createElement('div');
        bar.id = 'updatePrompt';
        bar.style.position = 'fixed';
        bar.style.bottom = '12px';
        bar.style.right = '12px';
        bar.style.zIndex = '500';
        bar.style.background = 'var(--panel)';
        bar.style.border = '1px solid var(--border)';
        bar.style.padding = '12px 16px';
        bar.style.borderRadius = '10px';
        bar.style.boxShadow = '0 6px 24px -6px rgba(0,0,0,0.5)';
        bar.innerHTML = `<span style='margin-right:12px'>Є оновлення. Перезавантажити?</span><button class='primary' id='reloadNowBtn'>Так</button><button class='subtle' id='reloadLaterBtn'>Пізніше</button>`;
        document.body.appendChild(bar);
        bar.querySelector('#reloadNowBtn').onclick = () => location.reload();
        bar.querySelector('#reloadLaterBtn').onclick = () => bar.remove();
      }
    }
  } catch (e) {
    // silent fail
  }
}
setTimeout(checkForUpdate, 15_000); // initial delayed check
setInterval(checkForUpdate, VERSION_CHECK_INTERVAL_MS);

window.addEventListener('keydown', (e) => {
  if (e.ctrlKey && e.key === 's') {
    e.preventDefault();
    persist();
    showToast('Збережено', 'success');
  }
});

(function upgradeLayout() {
  const needProps = { align: 'center', type: 'text', format: '', formula: '', isCustom: false };
  state.layout.columns.forEach((c) => {
    Object.keys(needProps).forEach((k) => {
      if (c[k] === undefined) c[k] = needProps[k];
    });
  });
  if (state.layout.groupSubtotalMergeUntil === undefined)
    state.layout.groupSubtotalMergeUntil = 'issued';
  if (state.layout.dateRow === undefined) state.layout.dateRow = true;
})();

(function migrateAddWithheld() {
  let changed = false;
  state.employees.forEach((e) => {
    if (e.withheld === undefined) {
      e.withheld = 0;
      changed = true;
    }
  });
  if (changed) persist();
})();

// ===== Essential utilities (migrated from assets/app.js) =====
// Global error handling
function safeToast(msg, type = 'info', timeout = 4000) {
  try {
    showToast(msg, type, timeout);
  } catch (_) {
    try {
      if (type === 'error') console.error(msg);
      alert(msg);
    } catch {}
  }
}
window.addEventListener('error', (ev) => {
  safeToast('Сталася помилка в застосунку. Перезавантажте сторінку.', 'error', 6000);
  if (ev?.error) console.error(ev.error);
});
window.addEventListener('unhandledrejection', (ev) => {
  safeToast('Невідпрацьована помилка. Перевірте введені дані.', 'error', 6000);
  console.error('Unhandled promise rejection:', ev?.reason);
});

// Offline/online notifications
window.addEventListener('offline', () => safeToast('Немає зʼєднання з інтернетом', 'warn'));
window.addEventListener('online', () => safeToast('Зʼєднання відновлено', 'success'));

// Auto-backup to localStorage
const STORAGE_KEYS = {
  EMP: 'payroll_employees_v1',
  META: 'payroll_meta_v1',
  LAY: 'payroll_excel_layout_v1',
  SET: 'payroll_settings_v1',
  BAK: 'payroll_autobackup_v1',
};
function snapshotAll() {
  return {
    ts: new Date().toISOString(),
    [STORAGE_KEYS.EMP]: localStorage.getItem(STORAGE_KEYS.EMP) || '[]',
    [STORAGE_KEYS.META]: localStorage.getItem(STORAGE_KEYS.META) || '{}',
    [STORAGE_KEYS.LAY]: localStorage.getItem(STORAGE_KEYS.LAY) || '{}',
    [STORAGE_KEYS.SET]: localStorage.getItem(STORAGE_KEYS.SET) || '{}',
  };
}
function saveAutoBackup() {
  try {
    localStorage.setItem(STORAGE_KEYS.BAK, JSON.stringify(snapshotAll()));
  } catch (e) {
    console.warn('Auto-backup failed', e);
  }
}
setInterval(saveAutoBackup, 60_000);
document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') saveAutoBackup();
});

function downloadText(text, name, type = 'application/json') {
  const blob = new Blob([text], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}

function exportAll(asDownload = false) {
  const pretty = JSON.stringify(snapshotAll(), null, 2);
  if (asDownload) {
    const d = new Date();
    const name = `payroll_backup_${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}.json`;
    downloadText(pretty, name);
    safeToast('Бекап завантажено', 'success');
  } else {
    navigator.clipboard?.writeText(pretty).then(
      () => safeToast('Стан скопійовано в буфер', 'success'),
      () => downloadText(pretty, 'payroll_backup.json')
    );
  }
}
function importAllFromString(json) {
  try {
    const snap = JSON.parse(json);
    if (!snap || typeof snap !== 'object') throw new Error('Invalid');
    if (snap[STORAGE_KEYS.EMP]) localStorage.setItem(STORAGE_KEYS.EMP, snap[STORAGE_KEYS.EMP]);
    if (snap[STORAGE_KEYS.META]) localStorage.setItem(STORAGE_KEYS.META, snap[STORAGE_KEYS.META]);
    if (snap[STORAGE_KEYS.LAY]) localStorage.setItem(STORAGE_KEYS.LAY, snap[STORAGE_KEYS.LAY]);
    if (snap[STORAGE_KEYS.SET]) localStorage.setItem(STORAGE_KEYS.SET, snap[STORAGE_KEYS.SET]);
    safeToast('Імпортовано. Перезавантажую...', 'success');
    setTimeout(() => location.reload(), 500);
  } catch {
    safeToast('Некоректний JSON для імпорту', 'error');
  }
}

// Hotkeys: Cmd/Ctrl+Shift+E (copy), +D (download), +I (import)
document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  if (!mod || !e.shiftKey) return;
  const k = e.key.toLowerCase();
  if (k === 'e') {
    e.preventDefault();
    exportAll(false);
  } else if (k === 'd') {
    e.preventDefault();
    exportAll(true);
  } else if (k === 'i') {
    e.preventDefault();
    const input = prompt('Вставте JSON бекапу для імпорту:');
    if (input) importAllFromString(input);
  }
});

// Removed scroll-to-top button per request
