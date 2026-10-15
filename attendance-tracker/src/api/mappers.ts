// Convert the snake_case JSON returned by the database functions into app types.
import type { AppData, AuditEntry, DayRecord, Employee } from '../lib/types';

/* eslint-disable @typescript-eslint/no-explicit-any */
export function mapEmployee(e: any): Employee {
  return { id: e.id, fullName: e.full_name, role: e.role, active: e.active ?? true };
}

export function mapRecord(r: any): DayRecord {
  return {
    employeeId: r.employee_id,
    date: r.date,
    status: r.status,
    startTime: r.start_time ?? null,
    endTime: r.end_time ?? null,
    notes: r.notes ?? null,
    leaveType: r.leave_type ?? null,
    leaveFullDay: r.leave_full_day ?? null,
    leaveStart: r.leave_start ?? null,
    leaveEnd: r.leave_end ?? null,
    leaveHours: r.leave_hours == null ? null : Number(r.leave_hours),
    updatedAt: r.updated_at ?? null,
  };
}

export function mapAppData(d: any): AppData {
  return {
    me: mapEmployee(d.me),
    settings: {
      year: d.settings.year,
      defaultStart: d.settings.default_start,
      defaultEnd: d.settings.default_end,
      eveEnd: d.settings.eve_end ?? '14:00',
      holidays: (d.holidays ?? []).map((h: any) => ({ date: h.date, name: h.name, kind: h.kind })),
    },
    employees: (d.employees ?? []).map(mapEmployee),
    records: (d.records ?? []).map(mapRecord),
    lastChanges: (d.last_changes ?? []).map((c: any) => ({
      employeeId: c.employee_id,
      changedAt: c.changed_at,
      changedByName: c.changed_by_name,
      action: c.action,
      recordDate: c.record_date ?? null,
    })),
  };
}

export function mapAudit(a: any): AuditEntry {
  return {
    id: a.id,
    changedAt: a.changed_at,
    changedByName: a.changed_by_name,
    employeeId: a.employee_id ?? null,
    employeeName: a.employee_name ?? null,
    entity: a.entity,
    action: a.action,
    recordDate: a.record_date ?? null,
    previousValue: a.previous_value ?? null,
    newValue: a.new_value ?? null,
  };
}
