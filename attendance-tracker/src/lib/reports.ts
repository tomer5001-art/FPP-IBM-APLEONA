// Report builders: produce plain tables that the UI renders and the
// exporters turn into CSV or PDF.
import { computeMonth, computeYear, fmt, summarize, type ComputedDay, type Totals } from './attendance';
import { DAY_NAMES, formatDate, MONTH_NAMES, monthOf } from './dates';
import { STATUS_LABELS, type DayRecord, type Employee, type Settings } from './types';

export interface ReportTable {
  title: string;
  subtitle: string;
  headers: string[];
  rows: (string | number)[][];
  totals: [string, string][];
}

export function dayStatusLabel(d: ComputedDay): string {
  return d.indicator === 'weekend' ? 'Weekend' : STATUS_LABELS[d.status];
}

export function monthlyTotalsList(t: Totals): [string, string][] {
  return [
    ['Total working days', String(t.daysWorked)],
    ['Total worked hours', fmt(t.workedHours)],
    ['Total vacation days', fmt(t.vacationDays)],
    ['Total sick leave days', fmt(t.sickDays)],
    ['Total vacation hours', fmt(t.vacationHours)],
    ['Total sick leave hours', fmt(t.sickHours)],
  ];
}

export const MONTHLY_HEADERS = [
  'Date', 'Day', 'Status', 'Start', 'End', 'Worked hours', 'Vacation hours', 'Sick leave hours', 'Vacation days', 'Notes',
];

export function buildMonthlyReport(employee: Employee, month: number, records: DayRecord[], settings: Settings) {
  const days = computeMonth(month, records, employee.id, settings);
  const totals = summarize(days);
  const table: ReportTable = {
    title: `Monthly attendance report - ${employee.fullName}`,
    subtitle: `${MONTH_NAMES[month - 1]} 2026`,
    headers: MONTHLY_HEADERS,
    rows: days.map((d) => [
      formatDate(d.date),
      DAY_NAMES[d.dayOfWeek],
      dayStatusLabel(d),
      d.startTime ?? '',
      d.endTime ?? '',
      fmt(d.workedHours),
      fmt(d.vacationHours),
      fmt(d.sickHours),
      fmt(d.vacationDays),
      d.notes ?? '',
    ]),
    totals: monthlyTotalsList(totals),
  };
  return { days, totals, table };
}

export const ANNUAL_HEADERS = [
  'Employee', 'Normal working days', 'Days worked', 'Worked hours', 'Vacation days', 'Vacation hours',
  'Sick leave days', 'Sick leave hours', 'Manual work days',
];

export function annualRow(name: string, t: Totals): (string | number)[] {
  return [name, t.normalWorkingDays, t.daysWorked, fmt(t.workedHours), fmt(t.vacationDays), fmt(t.vacationHours),
    fmt(t.sickDays), fmt(t.sickHours), t.manualWorkDays];
}

export function buildAnnualReport(employees: Employee[], records: DayRecord[], settings: Settings) {
  const perEmployee = employees.map((e) => {
    const days = computeYear(records, e.id, settings);
    const months = MONTH_NAMES.map((name, i) => ({ name, totals: summarize(days.filter((d) => monthOf(d.date) === i + 1)) }));
    return { employee: e, totals: summarize(days), months };
  });
  const summary: ReportTable = {
    title: 'Annual report 2026',
    subtitle: 'All figures for 1 January - 31 December 2026',
    headers: ANNUAL_HEADERS,
    rows: perEmployee.map((p) => annualRow(p.employee.fullName, p.totals)),
    totals: [],
  };
  const breakdowns: ReportTable[] = perEmployee.map((p) => ({
    title: `Annual report 2026 - ${p.employee.fullName}`,
    subtitle: 'Monthly breakdown',
    headers: ['Month', ...ANNUAL_HEADERS.slice(1)],
    rows: [...p.months.map((m) => annualRow(m.name, m.totals)), annualRow('Total 2026', p.totals)],
    totals: [],
  }));
  return { perEmployee, summary, breakdowns };
}
