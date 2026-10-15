import { HOLIDAYS_2026 } from '../src/lib/holidays2026';
import type { DayRecord, DayStatus, Settings } from '../src/lib/types';

/** Default hours, no holidays: isolates the basic rules. */
export const SETTINGS: Settings = { year: 2026, defaultStart: '08:00', defaultEnd: '16:30', eveEnd: '14:00', holidays: [] };

/** Default hours with the 2026 Israeli holidays. */
export const SETTINGS_IL: Settings = { ...SETTINGS, holidays: HOLIDAYS_2026 };
export const EMP = 'emp-1';

export function rec(date: string, status: DayStatus, extra: Partial<DayRecord> = {}): DayRecord {
  return {
    employeeId: EMP,
    date,
    status,
    startTime: null,
    endTime: null,
    notes: null,
    leaveType: null,
    leaveFullDay: null,
    leaveStart: null,
    leaveEnd: null,
    leaveHours: null,
    updatedAt: null,
    ...extra,
  };
}
