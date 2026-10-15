import type { DayRecord, DayStatus, Settings } from '../src/lib/types';

export const SETTINGS: Settings = { year: 2026, defaultStart: '08:00', defaultEnd: '16:30' };
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
