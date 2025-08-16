import { state } from './state.js';
import { round2, fixedPerDay } from './utils.js';

/**
 * Recalculate pay for all employees (idempotent)
 */
export function computePays() {
  for (let emp of state.employees) {
    let pay = 0;
    const salesVal = Number(emp.sales) || 0;
    const giftsVal = Number(emp.gifts) || 0;
    const salesNet = salesVal - giftsVal;
    switch (emp.rateType) {
      case 'waiter':
        pay = salesNet * (emp.waiterPercent / 100);
        // Apply guarantee only if enabled, salesNet < threshold, and there is some activity (hours or sales/gifts entered)
        const hasActivity = emp.hoursMinutes > 0 || salesVal > 0 || giftsVal > 0;
        if (emp.waiterMinGuarantee !== false && hasActivity && salesNet < 10000) {
          pay = 500;
          emp.min500Applied = true;
        } else {
          emp.min500Applied = false;
        }
        break;
      case 'hostess':
        pay = (emp.hoursMinutes / 60) * emp.hourlyRate + salesNet * (emp.hostessPercent / 100);
        break;
      case 'fixed':
        pay = fixedPerDay(emp);
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
