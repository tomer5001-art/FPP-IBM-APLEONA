// Date helpers. Dates are plain "YYYY-MM-DD" strings and all calculations use
// UTC so the day of week never shifts with the device time zone.

export const YEAR = 2026;

export const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
export const DAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
export const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

const pad = (n: number) => String(n).padStart(2, '0');

export function toIsoDate(year: number, month: number, day: number): string {
  return `${year}-${pad(month)}-${pad(day)}`;
}

function parts(date: string): [number, number, number] {
  const [y, m, d] = date.split('-').map(Number);
  return [y, m, d];
}

/** 0 = Sunday ... 6 = Saturday */
export function dayOfWeek(date: string): number {
  const [y, m, d] = parts(date);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay();
}

/** Israeli work week for this app: Sunday - Thursday. */
export function isNormalWorkday(date: string): boolean {
  return dayOfWeek(date) <= 4;
}

export function isWeekend(date: string): boolean {
  return dayOfWeek(date) >= 5;
}

export function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** All dates of a month (month is 1-12). */
export function monthDates(year: number, month: number): string[] {
  return Array.from({ length: daysInMonth(year, month) }, (_, i) => toIsoDate(year, month, i + 1));
}

export function yearDates(year: number = YEAR): string[] {
  return Array.from({ length: 12 }, (_, i) => monthDates(year, i + 1)).flat();
}

export function monthOf(date: string): number {
  return parts(date)[1];
}

export function isInYear(date: string, year: number = YEAR): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(date) && parts(date)[0] === year;
}

export function addDays(date: string, days: number): string {
  const [y, m, d] = parts(date);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  return toIsoDate(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
}

/** Local "today" as YYYY-MM-DD. */
export function todayIso(now: Date = new Date()): string {
  return toIsoDate(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

/** The month to show by default: the current month inside 2026, else Jan/Dec. */
export function defaultMonth(now: Date = new Date()): number {
  if (now.getFullYear() < YEAR) return 1;
  if (now.getFullYear() > YEAR) return 12;
  return now.getMonth() + 1;
}

export function formatDate(date: string): string {
  const [y, m, d] = parts(date);
  return `${pad(d)}/${pad(m)}/${y}`;
}

export function formatLongDate(date: string): string {
  const [, m, d] = parts(date);
  return `${DAY_NAMES[dayOfWeek(date)]}, ${d} ${MONTH_NAMES[m - 1]} ${parts(date)[0]}`;
}

export function formatDateTime(iso: string): string {
  const dt = new Date(iso);
  if (Number.isNaN(dt.getTime())) return iso;
  return `${pad(dt.getDate())}/${pad(dt.getMonth() + 1)}/${dt.getFullYear()} ${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
}
