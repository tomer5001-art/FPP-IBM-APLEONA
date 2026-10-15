import { useEffect, useState } from 'react';
import { api } from '../api';
import { describeError, useApp } from '../context/AppContext';
import { formatDate, formatDateTime } from '../lib/dates';
import { STATUS_LABELS, type AuditEntry, type DayStatus } from '../lib/types';

/* eslint-disable @typescript-eslint/no-explicit-any */
function describeValue(entry: AuditEntry, v: any): string {
  if (v == null) return entry.entity === 'day' ? 'Default' : '—';
  if (entry.entity === 'day') {
    const parts = [STATUS_LABELS[v.status as DayStatus] ?? v.status];
    if (v.start_time) parts.push(`${v.start_time}-${v.end_time}`);
    if (v.leave && !v.leave.full_day) parts.push(`${v.leave.type} ${v.leave.start_time}-${v.leave.end_time}`);
    if (v.notes) parts.push(`"${v.notes}"`);
    return parts.join(' · ');
  }
  if (entry.entity === 'settings') return `${v.default_start}-${v.default_end}`;
  const parts = [v.full_name, v.role, v.active ? 'active' : 'inactive'];
  if (v.pin_changed) parts.push('PIN changed');
  return parts.join(' · ');
}

/** Audit history: who changed what and when, with previous and new values. */
export function ChangeList({ employeeId, limit = 50 }: { employeeId: string | null; limit?: number }) {
  const { token, data } = useApp();
  const [entries, setEntries] = useState<AuditEntry[] | null>(null);
  const [error, setError] = useState('');

  useEffect(() => {
    api
      .getAuditLog(token, employeeId, limit)
      .then(setEntries)
      .catch((e) => setError(describeError(e)));
  }, [token, employeeId, limit, data]);

  if (error) return <p className="alert">{error}</p>;
  if (!entries) return <p className="muted">Loading…</p>;
  if (entries.length === 0) return <p className="muted">No changes recorded yet.</p>;
  return (
    <ul className="change-list">
      {entries.map((e) => (
        <li key={e.id}>
          <div className="change-meta">
            <strong>
              {e.entity === 'day' && e.recordDate ? `Day ${formatDate(e.recordDate)}` : e.entity === 'settings' ? 'Settings' : 'Employee'}
              {e.employeeName && e.entity !== 'settings' ? ` · ${e.employeeName}` : ''}
            </strong>
            <span className="muted small">
              {formatDateTime(e.changedAt)} by {e.changedByName}
            </span>
          </div>
          <div className="change-values small">
            <span className="muted">{describeValue(e, e.previousValue)}</span>
            <span aria-label="changed to"> → </span>
            <span>{describeValue(e, e.newValue)}</span>
          </div>
        </li>
      ))}
    </ul>
  );
}
