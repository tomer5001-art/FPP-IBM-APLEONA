import { useEffect, useMemo, useState } from 'react';
import { describeError, useApp } from '../context/AppContext';
import {
  errorMessage,
  fmt,
  fmtHours,
  isPartialStatus,
  normalDayHours,
  resolveDay,
  suggestWorkForLeave,
  validateDay,
  type ComputedDay,
} from '../lib/attendance';
import { dayOfWeek, formatLongDate, isNormalWorkday } from '../lib/dates';
import { STATUS_LABELS, type DayRecord, type DayStatus } from '../lib/types';
import { StatusBadge } from './StatusBadge';

const WORKDAY_STATUSES: DayStatus[] = ['present', 'vacation', 'sick', 'partial_vacation', 'partial_sick', 'manual', 'no_work'];
const WEEKEND_STATUSES: DayStatus[] = ['no_work', 'manual'];

/** Daily record: view and (when allowed) edit one employee-day. */
export function DayDialog({ employeeId, day, onClose }: { employeeId: string; day: ComputedDay; onClose(): void }) {
  const { data, isAdmin, employeeById, saveDay, resetDay } = useApp();
  const settings = data.settings;
  const employee = employeeById(employeeId);
  const workday = isNormalWorkday(day.date);
  const readOnly = dayOfWeek(day.date) === 6 && !isAdmin;

  const [status, setStatus] = useState<DayStatus>(day.status);
  const [start, setStart] = useState(day.startTime ?? '');
  const [end, setEnd] = useState(day.endTime ?? '');
  const [leaveStart, setLeaveStart] = useState(day.leaveStart ?? '');
  const [leaveEnd, setLeaveEnd] = useState(day.leaveEnd ?? '');
  const [notes, setNotes] = useState(day.notes ?? '');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [onClose]);

  function changeStatus(s: DayStatus) {
    setStatus(s);
    setError('');
    if ((s === 'present' || s === 'manual') && !start && !end) {
      setStart(settings.defaultStart);
      setEnd(settings.defaultEnd);
    }
    if (isPartialStatus(s) && !leaveStart) {
      setStart('');
      setEnd('');
    }
  }

  function changeLeave(ls: string, le: string) {
    setLeaveStart(ls);
    setLeaveEnd(le);
    if (ls && le && !start && !end) {
      const w = suggestWorkForLeave(ls, le, settings);
      setStart(w.start);
      setEnd(w.end);
    }
  }

  const input = {
    employeeId,
    date: day.date,
    status,
    startTime: start || null,
    endTime: end || null,
    leaveStart: leaveStart || null,
    leaveEnd: leaveEnd || null,
    notes: notes || null,
  };

  const preview = useMemo(() => {
    const rec: DayRecord = {
      employeeId,
      date: day.date,
      status,
      startTime: start || null,
      endTime: end || null,
      notes: null,
      leaveType: null,
      leaveFullDay: null,
      leaveStart: leaveStart || null,
      leaveEnd: leaveEnd || null,
      leaveHours: null,
      updatedAt: null,
    };
    return resolveDay(day.date, rec, settings);
  }, [employeeId, day.date, status, start, end, leaveStart, leaveEnd, settings]);

  async function onSave() {
    const err = validateDay(input, isAdmin);
    if (err) return setError(errorMessage(err));
    setBusy(true);
    try {
      await saveDay(input);
      onClose();
    } catch (e) {
      setError(describeError(e));
    } finally {
      setBusy(false);
    }
  }

  async function onReset() {
    setBusy(true);
    try {
      await resetDay(employeeId, day.date);
      onClose();
    } catch (e) {
      setError(describeError(e));
    } finally {
      setBusy(false);
    }
  }

  const statuses = workday ? WORKDAY_STATUSES : WEEKEND_STATUSES;
  const showWork = status === 'present' || status === 'manual' || isPartialStatus(status);
  const leaveKind = status === 'partial_vacation' ? 'Vacation' : 'Sick leave';

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" role="dialog" aria-modal="true" aria-label="Daily record" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <h2>{employee?.fullName}</h2>
            <p className="muted">{formatLongDate(day.date)}</p>
          </div>
          <button type="button" className="btn btn-icon" onClick={onClose} aria-label="Close">
            ✕
          </button>
        </div>

        <div className="record-summary">
          <div>
            <span className="muted">Status</span>
            <StatusBadge day={day} />
          </div>
          <div>
            <span className="muted">Start</span>
            <strong>{day.startTime ?? '—'}</strong>
          </div>
          <div>
            <span className="muted">End</span>
            <strong>{day.endTime ?? '—'}</strong>
          </div>
          <div>
            <span className="muted">Worked</span>
            <strong>{fmtHours(day.workedHours)}</strong>
          </div>
          <div>
            <span className="muted">Vacation</span>
            <strong>{fmtHours(day.vacationHours)}</strong>
          </div>
          <div>
            <span className="muted">Sick leave</span>
            <strong>{fmtHours(day.sickHours)}</strong>
          </div>
        </div>
        {day.notes && <p className="note-box">{day.notes}</p>}
        {!day.isException && (
          <p className="muted small">
            {workday ? 'Automatic working day (no report needed).' : 'Weekend - no automatic working hours.'}
          </p>
        )}

        {readOnly ? (
          <p className="alert">Only the administrator can record work on Saturday.</p>
        ) : (
          <form
            className="form"
            onSubmit={(e) => {
              e.preventDefault();
              onSave();
            }}
          >
            <h3>Edit record</h3>
            <label className="field">
              <span>Status</span>
              <select value={status} onChange={(e) => changeStatus(e.target.value as DayStatus)}>
                {statuses.map((s) => (
                  <option key={s} value={s}>
                    {STATUS_LABELS[s]}
                  </option>
                ))}
              </select>
            </label>

            {isPartialStatus(status) && (
              <div className="row">
                <label className="field">
                  <span>{leaveKind} from</span>
                  <input type="time" value={leaveStart} onChange={(e) => changeLeave(e.target.value, leaveEnd)} />
                </label>
                <label className="field">
                  <span>{leaveKind} to</span>
                  <input type="time" value={leaveEnd} onChange={(e) => changeLeave(leaveStart, e.target.value)} />
                </label>
              </div>
            )}

            {showWork && (
              <div className="row">
                <label className="field">
                  <span>Work start</span>
                  <input type="time" value={start} onChange={(e) => setStart(e.target.value)} />
                </label>
                <label className="field">
                  <span>Work end</span>
                  <input type="time" value={end} onChange={(e) => setEnd(e.target.value)} />
                </label>
              </div>
            )}

            <label className="field">
              <span>Notes (optional)</span>
              <input
                type="text"
                value={notes}
                maxLength={500}
                placeholder="e.g. Doctor appointment"
                onChange={(e) => setNotes(e.target.value)}
              />
            </label>

            <p className="preview">
              Worked {fmtHours(preview.workedHours)}
              {preview.vacationHours > 0 && ` · Vacation ${fmtHours(preview.vacationHours)} (${fmt(preview.vacationDays)} day)`}
              {preview.sickHours > 0 && ` · Sick ${fmtHours(preview.sickHours)} (${fmt(preview.sickDays)} day)`}
              <span className="muted"> · Normal day = {fmt(normalDayHours(settings))} h</span>
            </p>

            {error && <p className="alert">{error}</p>}

            <div className="actions">
              <button type="submit" className="btn btn-primary" disabled={busy}>
                Save
              </button>
              {day.isException && (
                <button type="button" className="btn" onClick={onReset} disabled={busy}>
                  Reset to default
                </button>
              )}
              <button type="button" className="btn btn-ghost" onClick={onClose}>
                Cancel
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}
