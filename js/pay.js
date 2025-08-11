import { state } from './state.js';
import { round2 } from './utils.js';

/**
 * Recalculate pay for all employees (idempotent)
 */
export function computePays() {
  for (let emp of state.employees) {
    let pay = 0;
    const salesNet = (Number(emp.sales) || 0) - (Number(emp.gifts) || 0);
    switch (emp.rateType) {
      case 'waiter':
        pay = salesNet * (emp.waiterPercent / 100);
        break;
      case 'hostess':
        pay = (emp.hoursMinutes / 60) * emp.hourlyRate + salesNet * (emp.hostessPercent / 100);
        break;
      case 'fixed':
        pay = emp.basePay;
        break;
      case 'hourly':
      default:
        pay = (emp.hoursMinutes / 60) * emp.hourlyRate;
    }
    pay -= Number(emp.withheld) || 0;
    emp.pay = round2(pay);
  }
}

export function isDayOff(emp) {
  return /^(в|вихід|вибув)/i.test(emp.hoursText || '');
}
