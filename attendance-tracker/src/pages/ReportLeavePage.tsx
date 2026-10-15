import { useMemo, useState } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { EmployeeSelect } from '../components/Selectors';
import { describeError, useApp } from '../context/AppContext';
import { useSelectedEmployee } from './CalendarPage';
import {
  errorMessage,
  fmt,
  hoursBetween,
  normalDayHours,
  suggestWorkForLeave,
  validateDay,
} from '../lib/attendance';
import { addDays, dayOfWeek, formatLongDate, isInYear, isNormalWorkday, todayIso } from '../lib/dates';
import type { DayStatus, SaveDayInput } from '../lib/types';

type ReportType = 'vacation' | 'sick' | 'partial_vacation' | 'partial_sick' | 'manual';

const TYPES: { id: ReportType; label: string; help: string }[] = [
  { id: 'vacation', label: 'Vacation', help: 'Full vacation days. Counts 1 vacation day per working day.' },
  { id: 'sick', label: 'Sick Leave', help: 'Full sick days. Counts 1 sick day per working day.' },
  { id: 'partial_vacation', label: 'Partial Vacation', help: 'Vacation for part of the day, e.g. 12:00-16:30.' },
  { id: 'partial_sick', label: 'Partial Sick Leave', help: 'Sick leave for part of the day, e.g. a doctor appointment.' },
  { id: 'manual', label: 'Manual Work Hours', help: 'Different working hours, or work on Friday.' },
];

function defaultDate(): string {
  const t = todayIso();
  return isInYear(t) ? t : '2026-01-01';
}

/** Working days (Sun-Thu) between two dates, inclusive. */
function workdaysBetween(from: string, to: string): string[] {
  const out: string[] = [];
  if (!isInYear(from) || !isInYear(to) || to < from) return out;
  for (let d = from; d <= to; d = addDays(d, 1)) if (isNormalWorkday(d)) out.push(d);
  return out;
}

export function ReportLeavePage() {
  const { data, isAdmin, saveDays } = useApp();
  const settings = data.settings;
  const navigate = useNavigate();
  const [params] = useSearchParams();
  const initialType = (TYPES.find((t) => t.id === params.get('type'))?.id ?? 'vacation') as ReportType;
  const [employeeId, setEmployeeId] = useSelectedEmployee();

  const [type, setType] = useState<ReportType>(initialType);
  const [from, setFrom] = useState(defaultDate);
  const [to, setTo] = useState(defaultDate);
  const [start, setStart] = useState(settings.defaultStart);
  const [end, setEnd] = useState(settings.defaultEnd);
  const [leaveStart, setLeaveStart] = useState('');
  const [leaveEnd, setLeaveEnd] = useState('');
  const [notes, setNotes] = useState('');
  const [error, setError] = useState('');
  const [done, setDone] = useState('');
  const [busy, setBusy] = useState(false);

  const isFull = type === 'vacation' || type === 'sick';
  const isPartial = type === 'partial_vacation' || type === 'partial_sick';
  const fullDays = useMemo(() => (isFull ? workdaysBetween(from, to) : []), [isFull, from, to]);
  const existing = fullDays.filter((d) => data.records.some((r) => r.employeeId === employeeId && r.date === d)).length;

  function pickType(t: ReportType) {
    setType(t);
    setError('');
    setDone('');
    if (t === 'manual') {
      setStart(settings.defaultStart);
      setEnd(settings.defaultEnd);
    } else if (t === 'partial_vacation' || t === 'partial_sick') {
      setStart('');
      setEnd('');
      setLeaveStart('');
      setLeaveEnd('');
    }
  }

  function changeLeave(ls: string, le: string) {
    setLeaveStart(ls);
    setLeaveEnd(le);
    if (ls && le) {
      const w = suggestWorkForLeave(ls, le, settings);
      setStart(w.start);
      setEnd(w.end);
    }
  }

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    setError('');
    setDone('');
    let inputs: SaveDayInput[];
    if (isFull) {
      if (fullDays.length === 0) return setError('There are no working days (Sunday - Thursday) in the selected dates.');
      inputs = fullDays.map((date) => ({ employeeId, date, status: type as DayStatus, notes }));
    } else {
      inputs = [
        {
          employeeId,
          date: from,
          status: type as DayStatus,
          startTime: start || null,
          endTime: end || null,
          leaveStart: isPartial ? leaveStart : null,
          leaveEnd: isPartial ? leaveEnd : null,
          notes,
        },
      ];
    }
    for (const i of inputs) {
      const err = validateDay(i, isAdmin);
      if (err) return setError(errorMessage(err));
    }
    setBusy(true);
    try {
      await saveDays(inputs);
      setDone(inputs.length === 1 ? 'Saved.' : `Saved ${inputs.length} days.`);
      setNotes('');
    } catch (err) {
      setError(describeError(err));
    } finally {
      setBusy(false);
    }
  }

  const leaveHours = hoursBetween(leaveStart, leaveEnd);
  const normal = normalDayHours(settings);
  const leaveWord = type === 'partial_vacation' ? 'Vacation' : 'Sick leave';

  return (
    <div className="stack">
      <div className="page-head">
        <h1>Report Leave</h1>
      </div>
      <form className="card form" onSubmit={submit}>
        {isAdmin && <EmployeeSelect employees={data.employees.filter((x) => x.active)} value={employeeId} onChange={setEmployeeId} />}

        <div className="type-grid" role="radiogroup" aria-label="Type">
          {TYPES.map((t) => (
            <button
              key={t.id}
              type="button"
              role="radio"
              aria-checked={type === t.id}
              className={`type-btn ${type === t.id ? 'is-active' : ''} type-${t.id}`}
              onClick={() => pickType(t.id)}
            >
              {t.label}
            </button>
          ))}
        </div>
        <p className="muted small">{TYPES.find((t) => t.id === type)?.help}</p>

        {isFull ? (
          <div className="row">
            <label className="field">
              <span>From</span>
              <input type="date" min="2026-01-01" max="2026-12-31" value={from} required
                onChange={(e) => { setFrom(e.target.value); if (e.target.value > to) setTo(e.target.value); }} />
            </label>
            <label className="field">
              <span>To</span>
              <input type="date" min={from} max="2026-12-31" value={to} required onChange={(e) => setTo(e.target.value)} />
            </label>
          </div>
        ) : (
          <label className="field">
            <span>Date</span>
            <input type="date" min="2026-01-01" max="2026-12-31" value={from} required onChange={(e) => setFrom(e.target.value)} />
          </label>
        )}
        {isInYear(from) && !isFull && <p className="muted small">{formatLongDate(from)}</p>}
        {isFull && (
          <p className="muted small">
            {fullDays.length} working day{fullDays.length === 1 ? '' : 's'} = {fullDays.length} {type === 'vacation' ? 'vacation' : 'sick'} day
            {fullDays.length === 1 ? '' : 's'}. Fridays and Saturdays are skipped.
            {existing > 0 && ` ${existing} existing entr${existing === 1 ? 'y' : 'ies'} will be replaced.`}
          </p>
        )}

        {isPartial && (
          <div className="row">
            <label className="field">
              <span>{leaveWord} from</span>
              <input type="time" value={leaveStart} required onChange={(e) => changeLeave(e.target.value, leaveEnd)} />
            </label>
            <label className="field">
              <span>{leaveWord} to</span>
              <input type="time" value={leaveEnd} required onChange={(e) => changeLeave(leaveStart, e.target.value)} />
            </label>
          </div>
        )}
        {isPartial && leaveHours > 0 && (
          <p className="preview">
            {fmt(leaveHours)} hours = {fmt(leaveHours / normal)} {type === 'partial_vacation' ? 'vacation' : 'sick'} day
          </p>
        )}

        {!isFull && (
          <div className="row">
            <label className="field">
              <span>{isPartial ? 'Worked from (optional)' : 'Start time'}</span>
              <input type="time" value={start} required={!isPartial} onChange={(e) => setStart(e.target.value)} />
            </label>
            <label className="field">
              <span>{isPartial ? 'Worked to (optional)' : 'End time'}</span>
              <input type="time" value={end} required={!isPartial} onChange={(e) => setEnd(e.target.value)} />
            </label>
          </div>
        )}
        {type === 'manual' && dayOfWeek(from) === 6 && !isAdmin && (
          <p className="alert">Only the administrator can record work on Saturday.</p>
        )}

        <label className="field">
          <span>Notes (optional)</span>
          <input type="text" value={notes} maxLength={500} placeholder="e.g. Annual vacation" onChange={(e) => setNotes(e.target.value)} />
        </label>

        {error && <p className="alert">{error}</p>}
        {done && <p className="success">{done}</p>}

        <div className="actions">
          <button type="submit" className="btn btn-primary btn-large" disabled={busy}>
            {busy ? 'Saving…' : 'Submit'}
          </button>
          <button type="button" className="btn btn-large btn-ghost" onClick={() => navigate(isAdmin ? `/calendar?employee=${employeeId}` : '/calendar')}>
            Open calendar
          </button>
        </div>
      </form>
    </div>
  );
}
