import type { Holiday } from './types';

/**
 * Israeli holidays 2026 (5786/5787), same list as the database seed.
 * Used by the demo backend; with Supabase the list comes from the database
 * and the admin can edit it in Settings.
 */
export const HOLIDAYS_2026: Holiday[] = [
  { date: '2026-04-01', name: 'Erev Pesach', kind: 'eve' },
  { date: '2026-04-02', name: 'Pesach', kind: 'holiday' },
  { date: '2026-04-07', name: 'Erev Seventh Day of Pesach', kind: 'eve' },
  { date: '2026-04-08', name: 'Seventh Day of Pesach', kind: 'holiday' },
  { date: '2026-04-22', name: 'Independence Day', kind: 'holiday' },
  { date: '2026-05-21', name: 'Erev Shavuot', kind: 'eve' },
  { date: '2026-05-22', name: 'Shavuot', kind: 'holiday' },
  { date: '2026-09-11', name: 'Erev Rosh Hashana', kind: 'eve' },
  { date: '2026-09-12', name: 'Rosh Hashana', kind: 'holiday' },
  { date: '2026-09-13', name: 'Rosh Hashana (second day)', kind: 'holiday' },
  { date: '2026-09-20', name: 'Erev Yom Kippur', kind: 'eve' },
  { date: '2026-09-21', name: 'Yom Kippur', kind: 'holiday' },
  { date: '2026-09-25', name: 'Erev Sukkot', kind: 'eve' },
  { date: '2026-09-26', name: 'Sukkot', kind: 'holiday' },
  { date: '2026-10-02', name: 'Erev Simchat Torah', kind: 'eve' },
  { date: '2026-10-03', name: 'Simchat Torah', kind: 'holiday' },
];
