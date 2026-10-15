import { describe, expect, it } from 'vitest';
import {
  computeMonth,
  computeYear,
  hoursBetween,
  normalDayHours,
  resolveDay,
  suggestWorkForLeave,
  summarize,
  validateDay,
} from '../src/lib/attendance';
import { dayOfWeek, isNormalWorkday, monthDates, yearDates } from '../src/lib/dates';
import { EMP, rec, SETTINGS, SETTINGS_IL } from './helpers';

describe('dates 2026', () => {
  it('knows every date of 2026', () => {
    const all = yearDates();
    expect(all).toHaveLength(365);
    expect(all[0]).toBe('2026-01-01');
    expect(all[364]).toBe('2026-12-31');
    expect(monthDates(2026, 2)).toHaveLength(28);
  });

  it('identifies days of the week correctly', () => {
    expect(dayOfWeek('2026-01-01')).toBe(4); // Thursday
    expect(dayOfWeek('2026-01-02')).toBe(5); // Friday
    expect(dayOfWeek('2026-01-03')).toBe(6); // Saturday
    expect(dayOfWeek('2026-01-04')).toBe(0); // Sunday
    expect(dayOfWeek('2026-03-01')).toBe(0); // Sunday
    expect(dayOfWeek('2026-10-15')).toBe(4); // Thursday
    expect(dayOfWeek('2026-12-31')).toBe(4); // Thursday
  });

  it('Sunday-Thursday are working days, Friday and Saturday are not', () => {
    expect(isNormalWorkday('2026-01-04')).toBe(true);
    expect(isNormalWorkday('2026-01-08')).toBe(true);
    expect(isNormalWorkday('2026-01-09')).toBe(false);
    expect(isNormalWorkday('2026-01-10')).toBe(false);
  });

  it('2026 has 261 Sunday-Thursday working days', () => {
    const days = computeYear([], EMP, SETTINGS);
    expect(summarize(days).normalWorkingDays).toBe(261);
  });
});

describe('automatic attendance', () => {
  it('normal day is 8.5 hours', () => {
    expect(normalDayHours(SETTINGS)).toBe(8.5);
  });

  it('Sunday-Thursday without a record is Present 08:00-16:30', () => {
    const d = resolveDay('2026-01-04', undefined, SETTINGS);
    expect(d.status).toBe('present');
    expect(d.startTime).toBe('08:00');
    expect(d.endTime).toBe('16:30');
    expect(d.workedHours).toBe(8.5);
    expect(d.isException).toBe(false);
  });

  it('Friday and Saturday have no automatic hours', () => {
    for (const date of ['2026-01-02', '2026-01-03']) {
      const d = resolveDay(date, undefined, SETTINGS);
      expect(d.status).toBe('no_work');
      expect(d.indicator).toBe('weekend');
      expect(d.workedHours).toBe(0);
      expect(d.startTime).toBeNull();
    }
  });

  it('January 2026: 21 working days, 178.5 hours', () => {
    const t = summarize(computeMonth(1, [], EMP, SETTINGS));
    expect(t.normalWorkingDays).toBe(21);
    expect(t.daysWorked).toBe(21);
    expect(t.workedHours).toBe(178.5);
  });

  it('uses the configured default hours', () => {
    const d = resolveDay('2026-01-04', undefined, { ...SETTINGS, defaultStart: '09:00', defaultEnd: '17:00' });
    expect(d.workedHours).toBe(8);
  });
});

describe('vacation', () => {
  it('full vacation day removes working hours and counts 1 day', () => {
    const d = resolveDay('2026-01-05', rec('2026-01-05', 'vacation'), SETTINGS);
    expect(d.workedHours).toBe(0);
    expect(d.startTime).toBeNull();
    expect(d.vacationDays).toBe(1);
    expect(d.vacationHours).toBe(8.5);
    expect(d.sickDays).toBe(0);
  });

  it('partial vacation 12:00-16:30 with work 08:00-12:00', () => {
    const d = resolveDay(
      '2026-01-05',
      rec('2026-01-05', 'partial_vacation', { startTime: '08:00', endTime: '12:00', leaveStart: '12:00', leaveEnd: '16:30' }),
      SETTINGS,
    );
    expect(d.workedHours).toBe(4);
    expect(d.vacationHours).toBe(4.5);
    expect(d.vacationDays).toBeCloseTo(4.5 / 8.5, 6);
    expect(d.indicator).toBe('partial');
  });

  it('4.25 hours of vacation equals 0.5 vacation day', () => {
    const d = resolveDay(
      '2026-01-05',
      rec('2026-01-05', 'partial_vacation', { startTime: '08:00', endTime: '12:15', leaveStart: '12:15', leaveEnd: '16:30' }),
      SETTINGS,
    );
    expect(d.vacationHours).toBe(4.25);
    expect(d.vacationDays).toBe(0.5);
    expect(d.workedHours).toBe(4.25);
  });
});

describe('sick leave', () => {
  it('full sick day is tracked separately from vacation', () => {
    const d = resolveDay('2026-02-01', rec('2026-02-01', 'sick'), SETTINGS);
    expect(d.sickDays).toBe(1);
    expect(d.sickHours).toBe(8.5);
    expect(d.vacationDays).toBe(0);
    expect(d.workedHours).toBe(0);
  });

  it('partial sick leave hours', () => {
    const d = resolveDay(
      '2026-02-01',
      rec('2026-02-01', 'partial_sick', { startTime: '10:00', endTime: '16:30', leaveStart: '08:00', leaveEnd: '10:00' }),
      SETTINGS,
    );
    expect(d.sickHours).toBe(2);
    expect(d.sickDays).toBeCloseTo(2 / 8.5, 6);
    expect(d.workedHours).toBe(6.5);
    expect(d.vacationHours).toBe(0);
  });
});

describe('manual work', () => {
  it('manual hours 07:30-15:30 = 8 hours', () => {
    const d = resolveDay('2026-01-06', rec('2026-01-06', 'manual', { startTime: '07:30', endTime: '15:30' }), SETTINGS);
    expect(d.workedHours).toBe(8);
    expect(d.indicator).toBe('manual');
  });

  it('Friday manual work counts as a worked day', () => {
    const records = [rec('2026-01-09', 'manual', { startTime: '08:00', endTime: '12:00' })];
    const d = resolveDay('2026-01-09', records[0], SETTINGS);
    expect(d.workedHours).toBe(4);
    const t = summarize(computeMonth(1, records, EMP, SETTINGS));
    expect(t.daysWorked).toBe(22);
    expect(t.workedHours).toBe(182.5);
    expect(t.manualWorkDays).toBe(1);
    expect(t.normalWorkingDays).toBe(21);
  });

  it('No Work on a working day removes the hours', () => {
    const d = resolveDay('2026-01-06', rec('2026-01-06', 'no_work'), SETTINGS);
    expect(d.workedHours).toBe(0);
    expect(d.indicator).toBe('no_work');
  });
});

describe('totals', () => {
  it('monthly totals combine all statuses', () => {
    const records = [
      rec('2026-03-01', 'vacation'),
      rec('2026-03-02', 'vacation'),
      rec('2026-03-03', 'partial_vacation', { startTime: '08:00', endTime: '12:15', leaveStart: '12:15', leaveEnd: '16:30' }),
      rec('2026-03-04', 'sick'),
      rec('2026-03-05', 'partial_sick', { startTime: '12:30', endTime: '16:30', leaveStart: '08:00', leaveEnd: '12:15' }),
      rec('2026-03-06', 'manual', { startTime: '09:00', endTime: '13:00' }), // Friday
      rec('2026-03-08', 'manual', { startTime: '07:30', endTime: '15:30' }),
    ];
    const t = summarize(computeMonth(3, records, EMP, SETTINGS));
    // March 2026 has 23 Sun-Thu days
    expect(t.normalWorkingDays).toBe(23);
    // worked: 23 - 3 full leave days = 20 days with hours on workdays, + Friday
    expect(t.daysWorked).toBe(21);
    expect(t.vacationDays).toBe(2.5);
    expect(t.vacationHours).toBe(21.25);
    expect(t.sickDays).toBe(1.5);
    expect(t.sickHours).toBe(12.75);
    expect(t.manualWorkDays).toBe(2);
    // 17 default days x 8.5 + 4.25 + 4 + 4 (Fri) + 8
    expect(t.workedHours).toBe(17 * 8.5 + 4.25 + 4 + 4 + 8);
  });

  it('annual vacation totals add up across months', () => {
    const records = [
      rec('2026-01-05', 'vacation'),
      rec('2026-04-12', 'vacation'),
      rec('2026-08-02', 'partial_vacation', { leaveStart: '12:15', leaveEnd: '16:30', startTime: '08:00', endTime: '12:15' }),
      rec('2026-12-31', 'vacation'),
      { ...rec('2026-06-01', 'vacation'), employeeId: 'someone-else' },
    ];
    const t = summarize(computeYear(records, EMP, SETTINGS));
    expect(t.vacationDays).toBe(3.5);
    expect(t.vacationHours).toBe(29.75);
    expect(t.daysWorked).toBe(261 - 3);
  });
});

describe('validation', () => {
  const base = { employeeId: EMP };
  it('rejects Saturday work for employees but allows admin', () => {
    const input = { ...base, date: '2026-01-03', status: 'manual' as const, startTime: '09:00', endTime: '12:00' };
    expect(validateDay(input, false, SETTINGS)).toBe('saturday_admin_only');
    expect(validateDay(input, true, SETTINGS)).toBeNull();
  });
  it('allows Friday manual work for employees', () => {
    expect(validateDay({ ...base, date: '2026-01-02', status: 'manual', startTime: '08:00', endTime: '12:00' }, false, SETTINGS)).toBeNull();
  });
  it('rejects leave on Friday', () => {
    expect(validateDay({ ...base, date: '2026-01-02', status: 'vacation' }, false, SETTINGS)).toBe('leave_requires_workday');
  });
  it('requires hours for manual work and partial leave', () => {
    expect(validateDay({ ...base, date: '2026-01-04', status: 'manual' }, false, SETTINGS)).toBe('work_time_required');
    expect(validateDay({ ...base, date: '2026-01-04', status: 'partial_vacation' }, false, SETTINGS)).toBe('leave_time_required');
  });
  it('rejects reversed and overlapping times', () => {
    expect(validateDay({ ...base, date: '2026-01-04', status: 'manual', startTime: '15:00', endTime: '08:00' }, false, SETTINGS)).toBe('work_time_invalid');
    expect(
      validateDay({ ...base, date: '2026-01-04', status: 'partial_sick', startTime: '08:00', endTime: '13:00', leaveStart: '12:00', leaveEnd: '16:30' }, false, SETTINGS),
    ).toBe('time_overlap');
  });
  it('rejects dates outside 2026', () => {
    expect(validateDay({ ...base, date: '2025-12-31', status: 'vacation' }, true, SETTINGS)).toBe('date_out_of_range');
  });
});

describe('helpers', () => {
  it('hoursBetween', () => {
    expect(hoursBetween('08:00', '16:30')).toBe(8.5);
    expect(hoursBetween('16:30', '08:00')).toBe(0);
    expect(hoursBetween(null, '08:00')).toBe(0);
  });
  it('suggests the working part of a partial day', () => {
    expect(suggestWorkForLeave('12:00', '16:30', SETTINGS)).toEqual({ start: '08:00', end: '12:00' });
    expect(suggestWorkForLeave('08:00', '10:00', SETTINGS)).toEqual({ start: '10:00', end: '16:30' });
  });
});

describe('holidays and holiday eves', () => {
  it('2026 has 256 working days after 5 weekday holidays', () => {
    const t = summarize(computeYear([], EMP, SETTINGS_IL));
    expect(t.normalWorkingDays).toBe(256);
    expect(t.holidayDays).toBe(5); // Pesach, 7th of Pesach, Independence Day, Rosh Hashana II, Yom Kippur
    expect(t.daysWorked).toBe(256);
  });

  it('a holiday has no automatic work', () => {
    const d = resolveDay('2026-04-02', undefined, SETTINGS_IL);
    expect(d.kind).toBe('holiday');
    expect(d.indicator).toBe('holiday');
    expect(d.holidayName).toBe('Pesach');
    expect(d.workedHours).toBe(0);
    expect(d.isNormalWorkday).toBe(false);
  });

  it('a holiday eve is worked 08:00-14:00', () => {
    const d = resolveDay('2026-04-01', undefined, SETTINGS_IL);
    expect(d.kind).toBe('eve');
    expect(d.status).toBe('present');
    expect(d.startTime).toBe('08:00');
    expect(d.endTime).toBe('14:00');
    expect(d.workedHours).toBe(6);
    expect(d.isNormalWorkday).toBe(true);
  });

  it('eve end time follows the setting', () => {
    expect(resolveDay('2026-09-20', undefined, { ...SETTINGS_IL, eveEnd: '13:00' }).workedHours).toBe(5);
  });

  it('full vacation on an eve counts half a vacation day (4.25 hours)', () => {
    const d = resolveDay('2026-04-01', rec('2026-04-01', 'vacation'), SETTINGS_IL);
    expect(d.vacationDays).toBe(0.5);
    expect(d.vacationHours).toBe(4.25);
    expect(d.workedHours).toBe(0);
  });

  it('full sick day on an eve is still 1 sick day of 6 hours', () => {
    const d = resolveDay('2026-04-01', rec('2026-04-01', 'sick'), SETTINGS_IL);
    expect(d.sickDays).toBe(1);
    expect(d.sickHours).toBe(6);
  });

  it('vacation range over an eve: 3 normal days + eve = 3.5 days', () => {
    const records = ['2026-09-16', '2026-09-17', '2026-09-20', '2026-09-22'].map((d) => rec(d, 'vacation'));
    const t = summarize(computeMonth(9, records, EMP, SETTINGS_IL));
    expect(t.vacationDays).toBe(3.5);
    expect(t.vacationHours).toBe(3 * 8.5 + 4.25);
  });

  it('partial vacation on an eve uses the 8.5 hour day for the fraction', () => {
    expect(suggestWorkForLeave('12:00', '14:00', SETTINGS_IL, '2026-04-01')).toEqual({ start: '08:00', end: '12:00' });
    const d = resolveDay(
      '2026-04-01',
      rec('2026-04-01', 'partial_vacation', { startTime: '08:00', endTime: '12:00', leaveStart: '12:00', leaveEnd: '14:00' }),
      SETTINGS_IL,
    );
    expect(d.workedHours).toBe(4);
    expect(d.vacationHours).toBe(2);
    expect(d.vacationDays).toBeCloseTo(2 / 8.5, 6);
  });

  it('April 2026: 19 working days, 156.5 hours', () => {
    const t = summarize(computeMonth(4, [], EMP, SETTINGS_IL));
    expect(t.normalWorkingDays).toBe(19);
    expect(t.holidayDays).toBe(3);
    expect(t.workedHours).toBe(17 * 8.5 + 2 * 6);
  });

  it('leave cannot be recorded on a holiday, manual work can', () => {
    expect(validateDay({ employeeId: EMP, date: '2026-09-21', status: 'vacation' }, false, SETTINGS_IL)).toBe('leave_requires_workday');
    expect(validateDay({ employeeId: EMP, date: '2026-09-21', status: 'manual', startTime: '09:00', endTime: '12:00' }, false, SETTINGS_IL)).toBeNull();
    expect(validateDay({ employeeId: EMP, date: '2026-09-20', status: 'vacation' }, false, SETTINGS_IL)).toBeNull();
    const d = resolveDay('2026-09-21', rec('2026-09-21', 'manual', { startTime: '09:00', endTime: '12:00' }), SETTINGS_IL);
    expect(d.workedHours).toBe(3);
  });

  it('Saturday holiday stays a non-working day', () => {
    const d = resolveDay('2026-09-12', undefined, SETTINGS_IL);
    expect(d.indicator).toBe('holiday');
    expect(d.holidayName).toBe('Rosh Hashana');
    expect(validateDay({ employeeId: EMP, date: '2026-09-12', status: 'manual', startTime: '09:00', endTime: '12:00' }, false, SETTINGS_IL)).toBe('saturday_admin_only');
  });
});
