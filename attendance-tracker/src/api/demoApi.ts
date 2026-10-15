// Demo backend that runs entirely in the browser (localStorage).
// Used when Supabase is not configured, so the app can be tried locally.
// It mirrors the permission rules of the database functions, but data stays
// on this one device - use Supabase for real use.
import type { Api } from './api';
import { ApiError } from './api';
import { hoursBetween, isPartialStatus, normalDayHours, validateDay } from '../lib/attendance';
import { isNormalWorkday } from '../lib/dates';
import type { AuditEntry, DayRecord, Employee, SaveDayInput, Settings } from '../lib/types';

export interface KeyValueStore {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

interface DemoEmployee extends Employee {
  pinHash: string;
  createdAt: number;
}

interface DemoDb {
  version: 1;
  employees: DemoEmployee[];
  settings: Settings;
  records: DayRecord[];
  audit: AuditEntry[];
  sessions: { token: string; employeeId: string; expiresAt: number }[];
}

const STORAGE_KEY = 'attendance-2026-demo-db';

async function sha256(text: string): Promise<string> {
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return Array.from(new Uint8Array(bytes), (b) => b.toString(16).padStart(2, '0')).join('');
}

const pinHash = (id: string, pin: string) => sha256(`${id}:${pin}`);

function randomToken(): string {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  return Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
}

function uuid(): string {
  return crypto.randomUUID();
}

async function seed(): Promise<DemoDb> {
  const people: [string, 'admin' | 'employee', string][] = [
    ['Tomer Cohen', 'admin', '0000'],
    ['Nir Ben Haim', 'employee', '2222'],
    ['Jonatan Ben Sudai', 'employee', '1111'],
  ];
  const employees: DemoEmployee[] = [];
  for (const [i, [fullName, role, pin]] of people.entries()) {
    const id = uuid();
    employees.push({ id, fullName, role, active: true, pinHash: await pinHash(id, pin), createdAt: i });
  }
  return {
    version: 1,
    employees,
    settings: { year: 2026, defaultStart: '08:00', defaultEnd: '16:30' },
    records: [],
    audit: [],
    sessions: [],
  };
}

function publicEmployee(e: DemoEmployee): Employee {
  return { id: e.id, fullName: e.fullName, role: e.role, active: e.active };
}

function dayState(r: DayRecord | undefined) {
  if (!r) return null;
  return {
    status: r.status,
    start_time: r.startTime,
    end_time: r.endTime,
    notes: r.notes,
    leave: r.leaveType
      ? { type: r.leaveType, full_day: r.leaveFullDay, start_time: r.leaveStart, end_time: r.leaveEnd, hours: r.leaveHours }
      : null,
  };
}

export function createDemoApi(store: KeyValueStore = window.localStorage): Api {
  let cache: DemoDb | null = null;

  async function load(): Promise<DemoDb> {
    if (cache) return cache;
    const raw = store.getItem(STORAGE_KEY);
    if (raw) {
      try {
        cache = JSON.parse(raw) as DemoDb;
        return cache;
      } catch {
        /* fall through and reseed */
      }
    }
    cache = await seed();
    save();
    return cache;
  }

  function save() {
    if (cache) store.setItem(STORAGE_KEY, JSON.stringify(cache));
  }

  async function me(token: string): Promise<{ db: DemoDb; me: DemoEmployee }> {
    const db = await load();
    const s = db.sessions.find((x) => x.token === token && x.expiresAt > Date.now());
    const e = s && db.employees.find((x) => x.id === s.employeeId && x.active);
    if (!e) throw new ApiError('not_authenticated');
    return { db, me: e };
  }

  function assertAccess(db: DemoDb, me: DemoEmployee, employeeId: string) {
    if (!employeeId) throw new ApiError('employee_required');
    if (me.role !== 'admin' && me.id !== employeeId) throw new ApiError('forbidden');
    if (!db.employees.some((e) => e.id === employeeId)) throw new ApiError('employee_not_found');
  }

  function audit(db: DemoDb, me: DemoEmployee, entry: Omit<AuditEntry, 'id' | 'changedAt' | 'changedByName' | 'employeeName'>) {
    db.audit.unshift({
      ...entry,
      id: db.audit.length + 1,
      changedAt: new Date().toISOString(),
      changedByName: me.fullName,
      employeeName: db.employees.find((e) => e.id === entry.employeeId)?.fullName ?? null,
    });
  }

  return {
    mode: 'demo',

    async listLoginEmployees() {
      const db = await load();
      return db.employees.filter((e) => e.active).sort((a, b) => a.createdAt - b.createdAt)
        .map((e) => ({ id: e.id, fullName: e.fullName }));
    },

    async login(employeeId, pin) {
      const db = await load();
      const e = db.employees.find((x) => x.id === employeeId && x.active);
      if (!e || (await pinHash(e.id, pin)) !== e.pinHash) throw new ApiError('invalid_credentials');
      const token = randomToken();
      db.sessions.push({ token, employeeId: e.id, expiresAt: Date.now() + 30 * 86400_000 });
      save();
      return { token, employee: publicEmployee(e) };
    },

    async logout(token) {
      const db = await load();
      db.sessions = db.sessions.filter((s) => s.token !== token);
      save();
    },

    async getAppData(token) {
      const { db, me: m } = await me(token);
      const admin = m.role === 'admin';
      const visible = (id: string | null) => admin || id === m.id;
      const lastChanges = db.employees.filter((e) => visible(e.id)).flatMap((e) => {
        const a = db.audit.find((x) => x.entity === 'day' && x.employeeId === e.id);
        return a ? [{ employeeId: e.id, changedAt: a.changedAt, changedByName: a.changedByName, action: a.action, recordDate: a.recordDate }] : [];
      });
      return {
        me: publicEmployee(m),
        settings: { ...db.settings },
        employees: db.employees.filter((e) => visible(e.id)).sort((a, b) => a.createdAt - b.createdAt).map(publicEmployee),
        records: db.records.filter((r) => visible(r.employeeId)).map((r) => ({ ...r })),
        lastChanges,
      };
    },

    async saveDay(token, input: SaveDayInput) {
      const { db, me: m } = await me(token);
      assertAccess(db, m, input.employeeId);
      const err = validateDay(input, m.role === 'admin');
      if (err) throw new ApiError(err);

      const s = input.status;
      const workday = isNormalWorkday(input.date);
      const keepsWork = s === 'present' || s === 'manual' || isPartialStatus(s);
      const start = keepsWork ? input.startTime || null : null;
      const end = keepsWork ? input.endTime || null : null;
      const ls = isPartialStatus(s) ? input.leaveStart || null : null;
      const le = isPartialStatus(s) ? input.leaveEnd || null : null;
      const notes = (input.notes ?? '').trim() || null;

      const idx = db.records.findIndex((r) => r.employeeId === input.employeeId && r.date === input.date);
      const prev = idx >= 0 ? db.records[idx] : undefined;
      const isDefault =
        !notes &&
        ((workday && s === 'present' && (!start || (start === db.settings.defaultStart && end === db.settings.defaultEnd))) ||
          (!workday && s === 'no_work'));

      let next: DayRecord | undefined;
      if (!isDefault) {
        const leaveType = s === 'vacation' || s === 'partial_vacation' ? 'vacation' : s === 'sick' || s === 'partial_sick' ? 'sick' : null;
        const full = s === 'vacation' || s === 'sick';
        next = {
          employeeId: input.employeeId,
          date: input.date,
          status: s,
          startTime: start,
          endTime: end,
          notes,
          leaveType,
          leaveFullDay: leaveType ? full : null,
          leaveStart: ls,
          leaveEnd: le,
          leaveHours: leaveType ? (full ? normalDayHours(db.settings) : hoursBetween(ls, le)) : null,
          updatedAt: new Date().toISOString(),
        };
      }
      if (idx >= 0) db.records.splice(idx, 1);
      if (next) db.records.push(next);

      const p = dayState(prev);
      const n = dayState(next);
      if (JSON.stringify(p) !== JSON.stringify(n)) {
        audit(db, m, { employeeId: input.employeeId, entity: 'day', action: n ? 'update' : 'reset', recordDate: input.date, previousValue: p, newValue: n });
      }
      save();
    },

    async resetDay(token, employeeId, date) {
      const { db, me: m } = await me(token);
      assertAccess(db, m, employeeId);
      if (new Date(`${date}T00:00:00Z`).getUTCDay() === 6 && m.role !== 'admin') throw new ApiError('saturday_admin_only');
      const idx = db.records.findIndex((r) => r.employeeId === employeeId && r.date === date);
      if (idx >= 0) {
        const prev = dayState(db.records[idx]);
        db.records.splice(idx, 1);
        audit(db, m, { employeeId, entity: 'day', action: 'reset', recordDate: date, previousValue: prev, newValue: null });
        save();
      }
    },

    async updateSettings(token, defaultStart, defaultEnd) {
      const { db, me: m } = await me(token);
      if (m.role !== 'admin') throw new ApiError('forbidden');
      if (hoursBetween(defaultStart, defaultEnd) <= 0) throw new ApiError('work_time_invalid');
      const prev = { default_start: db.settings.defaultStart, default_end: db.settings.defaultEnd };
      db.settings = { ...db.settings, defaultStart, defaultEnd };
      audit(db, m, { employeeId: null, entity: 'settings', action: 'update', recordDate: null, previousValue: prev, newValue: { default_start: defaultStart, default_end: defaultEnd } });
      save();
    },

    async saveEmployee(token, input) {
      const { db, me: m } = await me(token);
      if (m.role !== 'admin') throw new ApiError('forbidden');
      const name = input.fullName.trim();
      const pin = (input.pin ?? '').trim();
      if (!name || name.length > 100) throw new ApiError('invalid_name');
      if (input.role !== 'admin' && input.role !== 'employee') throw new ApiError('invalid_role');
      if (pin && !/^\d{4,8}$/.test(pin)) throw new ApiError('invalid_pin_format');
      if (db.employees.some((e) => e.id !== input.id && e.fullName.toLowerCase() === name.toLowerCase())) throw new ApiError('name_taken');

      if (!input.id) {
        if (!pin) throw new ApiError('pin_required');
        const id = uuid();
        db.employees.push({ id, fullName: name, role: input.role, active: input.active, pinHash: await pinHash(id, pin), createdAt: Date.now() });
        audit(db, m, { employeeId: id, entity: 'employee', action: 'create', recordDate: null, previousValue: null, newValue: { full_name: name, role: input.role, active: input.active } });
        save();
        return id;
      }
      const e = db.employees.find((x) => x.id === input.id);
      if (!e) throw new ApiError('employee_not_found');
      const others = db.employees.filter((x) => x.id !== e.id && x.role === 'admin' && x.active);
      if (others.length === 0 && (input.role !== 'admin' || !input.active)) throw new ApiError('last_admin');
      const prev = { full_name: e.fullName, role: e.role, active: e.active };
      e.fullName = name;
      e.role = input.role;
      e.active = input.active;
      if (pin) e.pinHash = await pinHash(e.id, pin);
      if (pin || !input.active) db.sessions = db.sessions.filter((s) => s.employeeId !== e.id || s.token === token);
      audit(db, m, { employeeId: e.id, entity: 'employee', action: 'update', recordDate: null, previousValue: prev, newValue: { full_name: name, role: input.role, active: input.active, pin_changed: !!pin } });
      save();
      return e.id;
    },

    async getAuditLog(token, employeeId, limit = 50) {
      const { db, me: m } = await me(token);
      let filter = employeeId;
      if (m.role !== 'admin') {
        if (filter && filter !== m.id) throw new ApiError('forbidden');
        filter = m.id;
      }
      return db.audit
        .filter((a) => (!filter || a.employeeId === filter) && (m.role === 'admin' || a.entity === 'day'))
        .slice(0, limit);
    },
  };
}
