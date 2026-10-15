import { useMemo } from 'react';
import { computeYear, summarize } from '../lib/attendance';
import { defaultMonth, monthOf, todayIso } from '../lib/dates';
import { useApp } from './AppContext';

/** Year-level numbers for one employee, used by dashboards. */
export function useEmployeeStats(employeeId: string) {
  const { data } = useApp();
  return useMemo(() => {
    const year = computeYear(data.records, employeeId, data.settings);
    const month = defaultMonth();
    const today = todayIso();
    const monthDays = year.filter((d) => monthOf(d.date) === month);
    return {
      year,
      yearTotals: summarize(year),
      month,
      monthTotals: summarize(monthDays),
      monthToDate: summarize(monthDays.filter((d) => d.date <= today)),
      lastChange: data.lastChanges.find((c) => c.employeeId === employeeId),
    };
  }, [data, employeeId]);
}
