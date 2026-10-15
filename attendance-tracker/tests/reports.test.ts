import { describe, expect, it } from 'vitest';
import { tablesToCsv } from '../src/lib/export';
import { buildAnnualReport, buildMonthlyReport } from '../src/lib/reports';
import type { Employee } from '../src/lib/types';
import { EMP, rec, SETTINGS, SETTINGS_IL } from './helpers';

const employee: Employee = { id: EMP, fullName: 'Nir Ben Haim', role: 'employee', active: true };

describe('monthly report', () => {
  const records = [
    rec('2026-01-05', 'vacation', { notes: 'Annual vacation' }),
    rec('2026-01-06', 'partial_vacation', { startTime: '08:00', endTime: '12:00', leaveStart: '12:00', leaveEnd: '16:30' }),
    rec('2026-01-07', 'sick', { notes: 'Flu' }),
    rec('2026-01-09', 'manual', { startTime: '08:00', endTime: '12:00' }),
  ];
  const r = buildMonthlyReport(employee, 1, records, SETTINGS);

  it('has one row per day with date, weekday and status', () => {
    expect(r.table.rows).toHaveLength(31);
    expect(r.table.rows[0].slice(0, 5)).toEqual(['01/01/2026', 'Thursday', 'Present', '08:00', '16:30']);
    expect(r.table.rows[1].slice(0, 3)).toEqual(['02/01/2026', 'Friday', 'Weekend']);
    expect(r.table.rows[4].slice(0, 3)).toEqual(['05/01/2026', 'Monday', 'Vacation']);
    expect(r.table.rows[4][9]).toBe('Annual vacation');
    expect(r.table.rows[8].slice(0, 6)).toEqual(['09/01/2026', 'Friday', 'Manual Work', '08:00', '12:00', '4']);
  });

  it('computes monthly totals', () => {
    expect(r.totals.daysWorked).toBe(21 - 2 + 1);
    expect(r.totals.workedHours).toBe(18 * 8.5 + 4 + 4);
    expect(r.totals.vacationDays).toBe(1.53);
    expect(r.totals.vacationHours).toBe(13);
    expect(r.totals.sickDays).toBe(1);
    expect(r.totals.sickHours).toBe(8.5);
    expect(Object.fromEntries(r.table.totals)).toEqual({
      'Total working days': '20',
      'Total worked hours': '161',
      'Total vacation days': '1.53',
      'Total sick leave days': '1',
      'Total vacation hours': '13',
      'Total sick leave hours': '8.5',
    });
  });

  it('exports CSV with headers and totals', () => {
    const csv = tablesToCsv([r.table]);
    expect(csv).toContain('Date,Day,Status,Start,End,Worked hours,Vacation hours,Sick leave hours,Vacation days,Notes');
    expect(csv).toContain('05/01/2026,Monday,Vacation,,,0,8.5,0,1,Annual vacation');
    expect(csv).toContain('Total vacation days,1.53');
  });
});

describe('annual report', () => {
  it('summarises the year per employee', () => {
    const other: Employee = { id: 'emp-2', fullName: 'Jonatan Ben Sudai', role: 'employee', active: true };
    const records = [rec('2026-01-05', 'vacation'), rec('2026-05-10', 'sick'), { ...rec('2026-02-01', 'vacation'), employeeId: 'emp-2' }];
    const r = buildAnnualReport([employee, other], records, SETTINGS);
    expect(r.summary.rows[0]).toEqual(['Nir Ben Haim', 261, 0, 259, '2201.5', '1', '8.5', '1', '8.5', 0]);
    expect(r.summary.rows[1][5]).toBe('1');
    expect(r.breakdowns[0].rows).toHaveLength(13);
    expect(r.breakdowns[0].rows[12][0]).toBe('Total 2026');
  });
});

describe('reports with holidays', () => {
  it('labels holidays and eves in the monthly report', () => {
    const r = buildMonthlyReport(employee, 4, [], SETTINGS_IL);
    expect(r.table.rows[0].slice(0, 5)).toEqual(['01/04/2026', 'Wednesday', 'Present (Erev Pesach)', '08:00', '14:00']);
    expect(r.table.rows[0][5]).toBe('6');
    expect(r.table.rows[1].slice(0, 3)).toEqual(['02/04/2026', 'Thursday', 'Holiday - Pesach']);
    expect(r.totals.daysWorked).toBe(19);
  });

  it('annual report shows holidays and fewer working days', () => {
    const r = buildAnnualReport([employee], [], SETTINGS_IL);
    expect(r.summary.rows[0].slice(0, 4)).toEqual(['Nir Ben Haim', 256, 5, 256]);
  });
});
