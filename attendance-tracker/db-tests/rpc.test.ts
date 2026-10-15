// Tests the real SQL migrations and RPC functions on Postgres.
// A temporary database is created and dropped for the run, so the
// TEST_DATABASE_URL database itself is never modified.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import pg from 'pg';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';

const baseUrl = process.env.TEST_DATABASE_URL;
const dbName = `attendance_test_${Date.now()}`;
let admin: pg.Client;
let db: pg.Client;
const ids: Record<string, string> = {};

/** Call an RPC as the public "anon" role, exactly like the browser does. */
async function rpc(fn: string, args: Record<string, unknown> = {}) {
  const keys = Object.keys(args);
  const sql = `select public.${fn}(${keys.map((k, i) => `${k} => $${i + 1}`).join(', ')}) as r`;
  await db.query('set role anon');
  try {
    return (await db.query(sql, keys.map((k) => args[k]))).rows[0].r;
  } finally {
    await db.query('reset role');
  }
}

async function rpcError(fn: string, args: Record<string, unknown>) {
  try {
    await rpc(fn, args);
  } catch (e) {
    return (e as Error).message;
  }
  return 'no error';
}

async function login(name: string, pin: string): Promise<string> {
  const r = await rpc('login', { p_employee_id: ids[name], p_pin: pin });
  if (!r.token) throw new Error(`login failed: ${JSON.stringify(r)}`);
  return r.token;
}

describe.skipIf(!baseUrl)('database RPC functions', () => {
  let tomer: string;
  let nir: string;
  let jon: string;

  beforeAll(async () => {
    admin = new pg.Client({ connectionString: baseUrl });
    await admin.connect();
    await admin.query(`create database ${dbName}`);
    const url = new URL(baseUrl!);
    url.pathname = `/${dbName}`;
    db = new pg.Client({ connectionString: url.toString() });
    await db.connect();
    // Supabase has these roles already; plain Postgres needs them for the grants.
    await db.query(`do $$ begin
      if not exists (select 1 from pg_roles where rolname = 'anon') then create role anon nologin; end if;
      if not exists (select 1 from pg_roles where rolname = 'authenticated') then create role authenticated nologin; end if;
    end $$`);
    await db.query('grant usage on schema public to anon, authenticated');
    const dir = join(__dirname, '..', 'supabase', 'migrations');
    for (const f of readdirSync(dir).filter((f) => f.endsWith('.sql')).sort()) {
      await db.query(readFileSync(join(dir, f), 'utf8'));
    }
    for (const e of await rpc('list_login_employees')) ids[e.full_name] = e.id;
    tomer = await login('Tomer Cohen', '0000');
    nir = await login('Nir Ben Haim', '2222');
    jon = await login('Jonatan Ben Sudai', '1111');
  });

  afterAll(async () => {
    await db?.end();
    await admin?.query(`drop database if exists ${dbName} with (force)`);
    await admin?.end();
  });

  describe('seed and security', () => {
    it('creates the three employees with hashed PINs', async () => {
      expect(Object.keys(ids)).toEqual(['Tomer Cohen', 'Nir Ben Haim', 'Jonatan Ben Sudai']);
      const { rows } = await db.query('select full_name, role, pin_hash from employees order by created_at');
      expect(rows.map((r) => r.role)).toEqual(['admin', 'employee', 'employee']);
      for (const r of rows) {
        expect(r.pin_hash).toMatch(/^\$2a\$/);
        expect(r.pin_hash).not.toContain('2222');
        expect(r.pin_hash).not.toContain('1111');
      }
    });

    it('the public role cannot read tables directly', async () => {
      await db.query('set role anon');
      try {
        for (const t of ['employees', 'attendance_records', 'leave_records', 'audit_log', 'sessions', 'app_settings', 'holidays']) {
          await expect(db.query(`select * from public.${t}`)).rejects.toThrow(/permission denied/);
        }
        await expect(db.query(`select public._current_employee('x')`)).rejects.toThrow(/permission denied/);
      } finally {
        await db.query('reset role');
      }
    });

    it('rejects wrong PINs and locks after 5 attempts', async () => {
      // Use a dedicated employee so the seeded logins stay usable.
      await rpc('save_employee', { p_token: tomer, p_id: null, p_full_name: 'Lock Test', p_role: 'employee', p_active: true, p_pin: '4444' });
      const id = (await rpc('list_login_employees')).find((e: { full_name: string }) => e.full_name === 'Lock Test').id;
      for (let i = 0; i < 4; i++) expect((await rpc('login', { p_employee_id: id, p_pin: '0000' })).error).toBe('invalid_credentials');
      expect((await rpc('login', { p_employee_id: id, p_pin: '0000' })).error).toBe('locked');
      expect((await rpc('login', { p_employee_id: id, p_pin: '4444' })).error).toBe('locked');
      await rpc('save_employee', { p_token: tomer, p_id: id, p_full_name: 'Lock Test', p_role: 'employee', p_active: false, p_pin: null });
    });

    it('rejects invalid tokens', async () => {
      expect(await rpcError('get_app_data', { p_token: 'x'.repeat(64) })).toBe('not_authenticated');
      expect(await rpcError('get_app_data', { p_token: null })).toBe('not_authenticated');
    });
  });

  describe('employee permissions', () => {
    it('employee sees only their own data', async () => {
      await rpc('save_day', { p_token: tomer, p_employee_id: ids['Jonatan Ben Sudai'], p_date: '2026-02-01', p_status: 'vacation' });
      const d = await rpc('get_app_data', { p_token: nir });
      expect(d.me.full_name).toBe('Nir Ben Haim');
      expect(d.employees.map((e: { full_name: string }) => e.full_name)).toEqual(['Nir Ben Haim']);
      expect(d.records.every((r: { employee_id: string }) => r.employee_id === ids['Nir Ben Haim'])).toBe(true);
    });

    it('employee cannot edit or reset another employee', async () => {
      expect(await rpcError('save_day', { p_token: nir, p_employee_id: ids['Jonatan Ben Sudai'], p_date: '2026-02-02', p_status: 'vacation' })).toBe('forbidden');
      expect(await rpcError('reset_day', { p_token: nir, p_employee_id: ids['Jonatan Ben Sudai'], p_date: '2026-02-01' })).toBe('forbidden');
      expect(await rpcError('save_day', { p_token: jon, p_employee_id: ids['Tomer Cohen'], p_date: '2026-02-02', p_status: 'sick' })).toBe('forbidden');
    });

    it('employee cannot read another employee audit log', async () => {
      expect(await rpcError('get_audit_log', { p_token: nir, p_employee_id: ids['Jonatan Ben Sudai'] })).toBe('forbidden');
      const own = await rpc('get_audit_log', { p_token: jon, p_employee_id: null });
      expect(own.every((a: { employee_id: string }) => a.employee_id === ids['Jonatan Ben Sudai'])).toBe(true);
    });

    it('employee cannot call admin functions', async () => {
      expect(await rpcError('update_settings', { p_token: nir, p_default_start: '07:00', p_default_end: '15:00' })).toBe('forbidden');
      expect(await rpcError('save_employee', { p_token: nir, p_id: ids['Nir Ben Haim'], p_full_name: 'Nir Ben Haim', p_role: 'admin', p_active: true, p_pin: null })).toBe('forbidden');
    });

    it('employee cannot record Saturday work, admin can', async () => {
      expect(await rpcError('save_day', { p_token: nir, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-01-03', p_status: 'manual', p_start: '09:00', p_end: '12:00' })).toBe('saturday_admin_only');
      await rpc('save_day', { p_token: tomer, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-01-03', p_status: 'manual', p_start: '09:00', p_end: '12:00' });
    });
  });

  describe('attendance records', () => {
    async function day(token: string, employee: string, date: string) {
      const d = await rpc('get_app_data', { p_token: token });
      return d.records.find((r: { employee_id: string; date: string }) => r.employee_id === ids[employee] && r.date === date);
    }

    it('full vacation day stores a leave record of 8.5 hours', async () => {
      await rpc('save_day', { p_token: nir, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-01-05', p_status: 'vacation', p_notes: 'Annual vacation' });
      const r = await day(nir, 'Nir Ben Haim', '2026-01-05');
      expect(r).toMatchObject({ status: 'vacation', start_time: null, leave_type: 'vacation', leave_full_day: true, notes: 'Annual vacation' });
      expect(Number(r.leave_hours)).toBe(8.5);
    });

    it('partial vacation 12:15-16:30 = 4.25 hours', async () => {
      await rpc('save_day', {
        p_token: nir, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-01-06', p_status: 'partial_vacation',
        p_start: '08:00', p_end: '12:15', p_leave_start: '12:15', p_leave_end: '16:30',
      });
      const r = await day(nir, 'Nir Ben Haim', '2026-01-06');
      expect(r).toMatchObject({ start_time: '08:00', end_time: '12:15', leave_start: '12:15', leave_end: '16:30', leave_full_day: false });
      expect(Number(r.leave_hours)).toBe(4.25);
    });

    it('sick leave is stored separately as sick', async () => {
      await rpc('save_day', { p_token: jon, p_employee_id: ids['Jonatan Ben Sudai'], p_date: '2026-01-07', p_status: 'sick' });
      await rpc('save_day', {
        p_token: jon, p_employee_id: ids['Jonatan Ben Sudai'], p_date: '2026-01-08', p_status: 'partial_sick',
        p_start: '10:00', p_end: '16:30', p_leave_start: '08:00', p_leave_end: '10:00', p_notes: 'Doctor appointment',
      });
      const { rows } = await db.query(
        `select leave_date::text, leave_type, full_day, hours::float from leave_records where employee_id = $1 order by leave_date`,
        [ids['Jonatan Ben Sudai']],
      );
      expect(rows).toEqual([
        { leave_date: '2026-01-07', leave_type: 'sick', full_day: true, hours: 8.5 },
        { leave_date: '2026-01-08', leave_type: 'sick', full_day: false, hours: 2 },
        { leave_date: '2026-02-01', leave_type: 'vacation', full_day: true, hours: 8.5 },
      ]);
    });

    it('employee can enter manual work on Friday', async () => {
      await rpc('save_day', { p_token: nir, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-01-09', p_status: 'manual', p_start: '08:00', p_end: '12:00' });
      expect(await day(nir, 'Nir Ben Haim', '2026-01-09')).toMatchObject({ status: 'manual', start_time: '08:00', end_time: '12:00' });
    });

    it('rejects invalid input', async () => {
      const base = { p_token: nir, p_employee_id: ids['Nir Ben Haim'] };
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-09', p_status: 'vacation' })).toBe('leave_requires_workday');
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-11', p_status: 'manual' })).toBe('work_time_required');
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-11', p_status: 'manual', p_start: '15:00', p_end: '08:00' })).toBe('work_time_invalid');
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-11', p_status: 'partial_sick', p_start: '08:00', p_end: '13:00', p_leave_start: '12:00', p_leave_end: '16:30' })).toBe('time_overlap');
      expect(await rpcError('save_day', { ...base, p_date: '2025-12-31', p_status: 'vacation' })).toBe('date_out_of_range');
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-11', p_status: 'holiday' })).toBe('invalid_status');
      expect(await rpcError('save_day', { ...base, p_date: '2026-01-11', p_status: 'manual', p_start: '8am', p_end: '12:00' })).toBe('invalid_time');
    });

    it('a default day stores no row; reset removes the exception', async () => {
      const base = { p_token: nir, p_employee_id: ids['Nir Ben Haim'] };
      await rpc('save_day', { ...base, p_date: '2026-01-12', p_status: 'present', p_start: '08:00', p_end: '16:30' });
      expect(await day(nir, 'Nir Ben Haim', '2026-01-12')).toBeUndefined();
      await rpc('save_day', { ...base, p_date: '2026-01-12', p_status: 'manual', p_start: '07:30', p_end: '15:30' });
      expect(await day(nir, 'Nir Ben Haim', '2026-01-12')).toMatchObject({ status: 'manual' });
      await rpc('reset_day', { ...base, p_date: '2026-01-12' });
      expect(await day(nir, 'Nir Ben Haim', '2026-01-12')).toBeUndefined();
    });

    it('changing a vacation day to manual work removes the leave record', async () => {
      const base = { p_token: tomer, p_employee_id: ids['Nir Ben Haim'] };
      await rpc('save_day', { ...base, p_date: '2026-01-13', p_status: 'vacation' });
      await rpc('save_day', { ...base, p_date: '2026-01-13', p_status: 'manual', p_start: '09:00', p_end: '17:00' });
      const { rows } = await db.query(`select 1 from leave_records where employee_id = $1 and leave_date = '2026-01-13'`, [ids['Nir Ben Haim']]);
      expect(rows).toHaveLength(0);
    });
  });

  describe('audit log', () => {
    it('stores who, when, previous and new value', async () => {
      const log = await rpc('get_audit_log', { p_token: tomer, p_employee_id: ids['Nir Ben Haim'], p_limit: 100 });
      const change = log.find((a: { record_date: string; new_value: { status: string } }) => a.record_date === '2026-01-13' && a.new_value?.status === 'manual');
      expect(change.changed_by_name).toBe('Tomer Cohen');
      expect(change.changed_at).toBeTruthy();
      expect(change.previous_value.status).toBe('vacation');
      expect(change.new_value).toMatchObject({ status: 'manual', start_time: '09:00', end_time: '17:00' });
      const reset = log.find((a: { record_date: string; action: string }) => a.record_date === '2026-01-12' && a.action === 'reset');
      expect(reset.previous_value.status).toBe('manual');
      expect(reset.new_value).toBeNull();
    });

    it('get_app_data reports the last change per employee', async () => {
      const d = await rpc('get_app_data', { p_token: tomer });
      expect(d.last_changes.length).toBeGreaterThanOrEqual(2);
      expect(d.employees).toHaveLength(4);
    });
  });

  describe('holidays', () => {
    it('seeds the 2026 holidays and the 14:00 eve end', async () => {
      const d = await rpc('get_app_data', { p_token: nir });
      expect(d.settings.eve_end).toBe('14:00');
      expect(d.holidays).toHaveLength(16);
      expect(d.holidays.find((h: { date: string }) => h.date === '2026-09-21')).toMatchObject({ name: 'Yom Kippur', kind: 'holiday' });
      expect(d.holidays.find((h: { date: string }) => h.date === '2026-09-20')).toMatchObject({ name: 'Erev Yom Kippur', kind: 'eve' });
    });

    it('leave is rejected on a holiday, manual work is allowed', async () => {
      const base = { p_token: nir, p_employee_id: ids['Nir Ben Haim'] };
      expect(await rpcError('save_day', { ...base, p_date: '2026-09-21', p_status: 'vacation' })).toBe('leave_requires_workday');
      await rpc('save_day', { ...base, p_date: '2026-09-21', p_status: 'manual', p_start: '09:00', p_end: '11:00' });
      await rpc('reset_day', { ...base, p_date: '2026-09-21' });
    });

    it('vacation on a holiday eve stores the short day hours', async () => {
      const base = { p_token: nir, p_employee_id: ids['Nir Ben Haim'] };
      await rpc('save_day', { ...base, p_date: '2026-09-20', p_status: 'vacation' });
      const { rows } = await db.query(`select hours::float from leave_records where employee_id = $1 and leave_date = '2026-09-20'`, [ids['Nir Ben Haim']]);
      expect(rows[0].hours).toBe(6);
      // The default eve day (08:00-14:00) is not stored.
      await rpc('save_day', { ...base, p_date: '2026-09-20', p_status: 'present', p_start: '08:00', p_end: '14:00' });
      const after = await db.query(`select 1 from attendance_records where employee_id = $1 and work_date = '2026-09-20'`, [ids['Nir Ben Haim']]);
      expect(after.rows).toHaveLength(0);
    });

    it('only the admin can add, change and remove holidays', async () => {
      expect(await rpcError('save_holiday', { p_token: nir, p_date: '2026-10-27', p_name: 'Election Day', p_kind: 'holiday' })).toBe('forbidden');
      expect(await rpcError('delete_holiday', { p_token: jon, p_date: '2026-09-21' })).toBe('forbidden');
      await rpc('save_holiday', { p_token: tomer, p_date: '2026-10-27', p_name: 'Election Day', p_kind: 'holiday' });
      expect(await rpcError('save_day', { p_token: nir, p_employee_id: ids['Nir Ben Haim'], p_date: '2026-10-27', p_status: 'sick' })).toBe('leave_requires_workday');
      expect(await rpcError('save_holiday', { p_token: tomer, p_date: '2026-10-27', p_name: 'X', p_kind: 'party' })).toBe('invalid_kind');
      await rpc('delete_holiday', { p_token: tomer, p_date: '2026-10-27' });
      expect((await rpc('get_app_data', { p_token: nir })).holidays).toHaveLength(16);
      const log = await rpc('get_audit_log', { p_token: tomer, p_employee_id: null, p_limit: 10 });
      expect(log[0]).toMatchObject({ entity: 'holiday', action: 'delete', record_date: '2026-10-27', changed_by_name: 'Tomer Cohen' });
      expect(log[1]).toMatchObject({ entity: 'holiday', action: 'create' });
    });

    it('admin can change the eve end time', async () => {
      await rpc('update_settings', { p_token: tomer, p_default_start: '08:00', p_default_end: '16:30', p_eve_end: '13:00' });
      expect((await rpc('get_app_data', { p_token: nir })).settings.eve_end).toBe('13:00');
      expect(await rpcError('update_settings', { p_token: tomer, p_default_start: '08:00', p_default_end: '16:30', p_eve_end: '07:00' })).toBe('work_time_invalid');
      await rpc('update_settings', { p_token: tomer, p_default_start: '08:00', p_default_end: '16:30', p_eve_end: '14:00' });
    });
  });

  describe('admin functions', () => {
    it('admin can change default hours', async () => {
      await rpc('update_settings', { p_token: tomer, p_default_start: '07:30', p_default_end: '16:00' });
      expect((await rpc('get_app_data', { p_token: nir })).settings).toMatchObject({ default_start: '07:30', default_end: '16:00' });
      expect(await rpcError('update_settings', { p_token: tomer, p_default_start: '16:00', p_default_end: '08:00' })).toBe('work_time_invalid');
      await rpc('update_settings', { p_token: tomer, p_default_start: '08:00', p_default_end: '16:30' });
    });

    it('admin can add an employee without code changes', async () => {
      const r = await rpc('save_employee', { p_token: tomer, p_id: null, p_full_name: 'New Person', p_role: 'employee', p_active: true, p_pin: '5678' });
      ids['New Person'] = r.id;
      const token = await login('New Person', '5678');
      expect((await rpc('get_app_data', { p_token: token })).me.full_name).toBe('New Person');
      expect(await rpcError('save_employee', { p_token: tomer, p_id: null, p_full_name: 'new person', p_role: 'employee', p_active: true, p_pin: '1234' })).toBe('name_taken');
      expect(await rpcError('save_employee', { p_token: tomer, p_id: null, p_full_name: 'X', p_role: 'employee', p_active: true, p_pin: '12' })).toBe('invalid_pin_format');
    });

    it('changing a PIN signs the employee out and the new PIN works', async () => {
      const old = await login('New Person', '5678');
      await rpc('save_employee', { p_token: tomer, p_id: ids['New Person'], p_full_name: 'New Person', p_role: 'employee', p_active: true, p_pin: '9876' });
      expect(await rpcError('get_app_data', { p_token: old })).toBe('not_authenticated');
      expect((await rpc('login', { p_employee_id: ids['New Person'], p_pin: '5678' })).error).toBe('invalid_credentials');
      await login('New Person', '9876');
    });

    it('deactivating an employee blocks login and ends sessions', async () => {
      const t = await login('New Person', '9876');
      await rpc('save_employee', { p_token: tomer, p_id: ids['New Person'], p_full_name: 'New Person', p_role: 'employee', p_active: false, p_pin: null });
      expect(await rpcError('get_app_data', { p_token: t })).toBe('not_authenticated');
      expect((await rpc('login', { p_employee_id: ids['New Person'], p_pin: '9876' })).error).toBe('invalid_credentials');
      expect((await rpc('list_login_employees')).some((e: { full_name: string }) => e.full_name === 'New Person')).toBe(false);
    });

    it('the last admin cannot be demoted or deactivated', async () => {
      expect(await rpcError('save_employee', { p_token: tomer, p_id: ids['Tomer Cohen'], p_full_name: 'Tomer Cohen', p_role: 'employee', p_active: true, p_pin: null })).toBe('last_admin');
      expect(await rpcError('save_employee', { p_token: tomer, p_id: ids['Tomer Cohen'], p_full_name: 'Tomer Cohen', p_role: 'admin', p_active: false, p_pin: null })).toBe('last_admin');
      expect((await rpc('get_app_data', { p_token: tomer })).me.role).toBe('admin');
    });

    it('logout ends the session', async () => {
      const t = await login('Jonatan Ben Sudai', '1111');
      await rpc('logout', { p_token: t });
      expect(await rpcError('get_app_data', { p_token: t })).toBe('not_authenticated');
    });
  });
});
