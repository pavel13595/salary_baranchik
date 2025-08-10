// Entry point - wires modules
import { loadState, persist, state } from './state.js';
import { computePays } from './pay.js';
import { renderEmployeesTable, bindGlobalEvents, showToast } from './ui.js';

loadState();
computePays();
renderEmployeesTable();
bindGlobalEvents();

window.addEventListener('keydown', e => { if (e.ctrlKey && e.key === 's') { e.preventDefault(); persist(); showToast('Збережено', 'success'); } });

(function upgradeLayout() {
  const needProps = { align: 'center', type: 'text', format: '', formula: '', isCustom: false };
  state.layout.columns.forEach(c => { Object.keys(needProps).forEach(k => { if (c[k] === undefined) c[k] = needProps[k]; }); });
  if (state.layout.groupSubtotalMergeUntil === undefined) state.layout.groupSubtotalMergeUntil = 'issued';
  if (state.layout.dateRow === undefined) state.layout.dateRow = true;
})();

(function migrateAddWithheld() { let changed = false; state.employees.forEach(e => { if (e.withheld === undefined) { e.withheld = 0; changed = true; } }); if (changed) persist(); })();
