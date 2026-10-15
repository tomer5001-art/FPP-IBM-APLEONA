// Attendance rules and calculations. This is the single place that decides
// what a day "is" and how many hours / leave days it counts for.
import { dayOfWeek, isNormalWorkday, monthDates, yearDates, YEAR } from './dates';
import type { DayRecord, DayStatus, SaveDayInput, Settings } from './types';

export type Indicator = 'present' | 'vacation' | 'sick' | 'partial' | 'manual' | 'no_work' | 'weekend';

export interface ComputedDay {
  date: string;
  dayOfWeek: number;
  isNormalWorkday: boolean;
  isException: boolean; // a stored record exists for this day
  status: DayStatus;
  indicator: Indicator;
  startTime: string | null;
  endTime: string | null;
  leaveStart: string | null;
  leaveEnd: string | null;
  workedHours: number;
  vacationHours: number;
  sickHours: number;
  vacationDays: number; // full day = 1, partial = hours / normal day hours
  sickDays: number;
  notes: string | null;
}

export interface Totals {
  normalWorkingDays: number;
  daysWorked: number;
  workedHours: number;
  vacationDays: number;
  vacationHours: number;
  sickDays: number;
  sickHours: number;
  manualWorkDays: number;
}

/** Minutes since midnight for "HH:MM". */
export function toMinutes(time: string): number {
  const [h, m] = time.split(':').map(Number);
  return h * 60 + m;
}

/** Hours between two "HH:MM" values (0 if missing or reversed). */
export function hoursBetween(start: string | null | undefined, end: string | null | undefined): number {
  if (!start || !end) return 0;
  const diff = toMinutes(end) - toMinutes(start);
  return diff > 0 ? diff / 60 : 0;
}

export function normalDayHours(settings: Settings): number {
  return hoursBetween(settings.defaultStart, settings.defaultEnd);
}

export const round2 = (n: number) => Math.round(n * 100) / 100;

export function isLeaveStatus(status: DayStatus): boolean {
  return status === 'vacation' || status === 'sick' || status === 'partial_vacation' || status === 'partial_sick';
}

export function isPartialStatus(status: DayStatus): boolean {
  return status === 'partial_vacation' || status === 'partial_sick';
}

function indicatorFor(status: DayStatus, workday: boolean): Indicator {
  switch (status) {
    case 'present':
      return 'present';
    case 'vacation':
      return 'vacation';
    case 'sick':
      return 'sick';
    case 'partial_vacation':
    case 'partial_sick':
      return 'partial';
    case 'manual':
      return 'manual';
    case 'no_work':
      return workday ? 'no_work' : 'weekend';
  }
}

/** Resolve one day for one employee from the default rules + optional exception. */
export function resolveDay(date: string, record: DayRecord | undefined, settings: Settings): ComputedDay {
  const dow = dayOfWeek(date);
  const workday = isNormalWorkday(date);
  const normal = normalDayHours(settings);

  const day: ComputedDay = {
    date,
    dayOfWeek: dow,
    isNormalWorkday: workday,
    isException: !!record,
    status: workday ? 'present' : 'no_work',
    indicator: workday ? 'present' : 'weekend',
    startTime: workday ? settings.defaultStart : null,
    endTime: workday ? settings.defaultEnd : null,
    leaveStart: null,
    leaveEnd: null,
    workedHours: workday ? normal : 0,
    vacationHours: 0,
    sickHours: 0,
    vacationDays: 0,
    sickDays: 0,
    notes: null,
  };
  if (!record) return day;

  const status = record.status;
  day.status = status;
  day.indicator = indicatorFor(status, workday);
  day.notes = record.notes;
  day.startTime = null;
  day.endTime = null;
  day.workedHours = 0;

  switch (status) {
    case 'present':
      day.startTime = record.startTime ?? settings.defaultStart;
      day.endTime = record.endTime ?? settings.defaultEnd;
      day.workedHours = hoursBetween(day.startTime, day.endTime);
      break;
    case 'manual':
      day.startTime = record.startTime;
      day.endTime = record.endTime;
      day.workedHours = hoursBetween(day.startTime, day.endTime);
      break;
    case 'vacation':
      day.vacationHours = normal;
      day.vacationDays = 1;
      break;
    case 'sick':
      day.sickHours = normal;
      day.sickDays = 1;
      break;
    case 'partial_vacation':
    case 'partial_sick': {
      day.startTime = record.startTime;
      day.endTime = record.endTime;
      day.workedHours = hoursBetween(record.startTime, record.endTime);
      day.leaveStart = record.leaveStart;
      day.leaveEnd = record.leaveEnd;
      const hours = hoursBetween(record.leaveStart, record.leaveEnd);
      const fraction = normal > 0 ? hours / normal : 0;
      if (status === 'partial_vacation') {
        day.vacationHours = hours;
        day.vacationDays = fraction;
      } else {
        day.sickHours = hours;
        day.sickDays = fraction;
      }
      break;
    }
    case 'no_work':
      break;
  }
  return day;
}

export function indexRecords(records: DayRecord[], employeeId: string): Map<string, DayRecord> {
  const map = new Map<string, DayRecord>();
  for (const r of records) if (r.employeeId === employeeId) map.set(r.date, r);
  return map;
}

export function computeDays(dates: string[], records: DayRecord[], employeeId: string, settings: Settings): ComputedDay[] {
  const map = indexRecords(records, employeeId);
  return dates.map((d) => resolveDay(d, map.get(d), settings));
}

export function computeMonth(month: number, records: DayRecord[], employeeId: string, settings: Settings, year = YEAR) {
  return computeDays(monthDates(year, month), records, employeeId, settings);
}

export function computeYear(records: DayRecord[], employeeId: string, settings: Settings, year = YEAR) {
  return computeDays(yearDates(year), records, employeeId, settings);
}

export function summarize(days: ComputedDay[]): Totals {
  const t: Totals = {
    normalWorkingDays: 0,
    daysWorked: 0,
    workedHours: 0,
    vacationDays: 0,
    vacationHours: 0,
    sickDays: 0,
    sickHours: 0,
    manualWorkDays: 0,
  };
  for (const d of days) {
    if (d.isNormalWorkday) t.normalWorkingDays++;
    if (d.workedHours > 0) t.daysWorked++;
    if (d.status === 'manual') t.manualWorkDays++;
    t.workedHours += d.workedHours;
    t.vacationDays += d.vacationDays;
    t.vacationHours += d.vacationHours;
    t.sickDays += d.sickDays;
    t.sickHours += d.sickHours;
  }
  t.workedHours = round2(t.workedHours);
  t.vacationDays = round2(t.vacationDays);
  t.vacationHours = round2(t.vacationHours);
  t.sickDays = round2(t.sickDays);
  t.sickHours = round2(t.sickHours);
  return t;
}

/** "8.5" / "4.25" / "0" - compact number for tables. */
export function fmt(n: number): string {
  return String(round2(n));
}

/** "8:30" style hours display. */
export function fmtHours(n: number): string {
  const totalMinutes = Math.round(n * 60);
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${h}:${String(m).padStart(2, '0')}`;
}

// ---------------------------------------------------------------------------
// Validation. The server enforces the same rules; this gives instant feedback
// in the form and is reused by the demo backend.

export const ERROR_MESSAGES: Record<string, string> = {
  not_authenticated: 'Your session has expired. Please sign in again.',
  forbidden: 'You do not have permission to do that.',
  invalid_credentials: 'Incorrect PIN.',
  locked: 'Too many wrong PINs. Try again in 5 minutes.',
  employee_required: 'Please select an employee.',
  employee_not_found: 'Employee not found.',
  date_out_of_range: 'Only dates in 2026 can be recorded.',
  saturday_admin_only: 'Only the administrator can record work on Saturday.',
  invalid_status: 'Please choose a valid status.',
  invalid_time: 'Please enter times as HH:MM.',
  notes_too_long: 'Notes can be up to 500 characters.',
  work_time_incomplete: 'Please enter both a start and an end time.',
  work_time_invalid: 'The end time must be after the start time.',
  work_time_required: 'Please enter the working hours.',
  leave_time_incomplete: 'Please enter both leave start and end times.',
  leave_time_invalid: 'The leave end time must be after the leave start time.',
  leave_time_required: 'Please enter the leave hours.',
  leave_requires_workday: 'Leave can only be recorded on a normal working day (Sunday - Thursday).',
  time_overlap: 'Working hours and leave hours overlap.',
  invalid_name: 'Please enter a name.',
  invalid_role: 'Please choose a role.',
  invalid_pin_format: 'The PIN must be 4 to 8 digits.',
  pin_required: 'Please enter a PIN for the new employee.',
  name_taken: 'An employee with this name already exists.',
  last_admin: 'There must be at least one active administrator.',
};

export function errorMessage(code: string): string {
  return ERROR_MESSAGES[code] ?? code;
}

const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

/** Returns null when valid, otherwise an error code. Normalises the input in place of the server. */
export function validateDay(input: SaveDayInput, isAdmin: boolean): string | null {
  const dow = dayOfWeek(input.date);
  const workday = dow <= 4;
  if (!/^2026-\d{2}-\d{2}$/.test(input.date)) return 'date_out_of_range';
  if (dow === 6 && !isAdmin) return 'saturday_admin_only';
  const s = input.status;
  for (const t of [input.startTime, input.endTime, input.leaveStart, input.leaveEnd]) {
    if (t && !TIME_RE.test(t)) return 'invalid_time';
  }
  if ((input.notes ?? '').trim().length > 500) return 'notes_too_long';

  const needsWork = s === 'present' || s === 'manual' || isPartialStatus(s);
  const start = needsWork ? input.startTime || null : null;
  const end = needsWork ? input.endTime || null : null;
  const ls = isPartialStatus(s) ? input.leaveStart || null : null;
  const le = isPartialStatus(s) ? input.leaveEnd || null : null;

  if (!!start !== !!end) return 'work_time_incomplete';
  if (start && end && toMinutes(end) <= toMinutes(start)) return 'work_time_invalid';
  if (!!ls !== !!le) return 'leave_time_incomplete';
  if (ls && le && toMinutes(le) <= toMinutes(ls)) return 'leave_time_invalid';
  if (s === 'manual' && !start) return 'work_time_required';
  if (s === 'present' && !workday && !start) return 'work_time_required';
  if (isLeaveStatus(s) && !workday) return 'leave_requires_workday';
  if (isPartialStatus(s)) {
    if (!ls || !le) return 'leave_time_required';
    if (start && end && toMinutes(start) < toMinutes(le) && toMinutes(ls) < toMinutes(end)) return 'time_overlap';
  }
  return null;
}

/** For a partial leave, suggest the working part of the default day. */
export function suggestWorkForLeave(leaveStart: string, leaveEnd: string, settings: Settings) {
  const ds = toMinutes(settings.defaultStart);
  const de = toMinutes(settings.defaultEnd);
  const ls = toMinutes(leaveStart);
  const le = toMinutes(leaveEnd);
  if (ls > ds && le >= de) return { start: settings.defaultStart, end: leaveStart };
  if (ls <= ds && le < de) return { start: leaveEnd, end: settings.defaultEnd };
  return { start: '', end: '' };
}
