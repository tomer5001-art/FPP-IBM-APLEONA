import { useState } from 'react';
import { Navigate } from 'react-router-dom';
import { ChangeList } from '../components/ChangeList';
import { describeError, useApp } from '../context/AppContext';
import { fmt, hoursBetween } from '../lib/attendance';
import { DAY_SHORT, dayOfWeek, formatDate } from '../lib/dates';
import type { Employee, Holiday, Role } from '../lib/types';

function WorkHoursForm() {
  const { data, updateSettings } = useApp();
  const [start, setStart] = useState(data.settings.defaultStart);
  const [end, setEnd] = useState(data.settings.defaultEnd);
  const [eveEnd, setEveEnd] = useState(data.settings.eveEnd);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);

  async function save(e: React.FormEvent) {
    e.preventDefault();
    try {
      await updateSettings(start, end, eveEnd);
      setMsg({ ok: true, text: 'Saved.' });
    } catch (err) {
      setMsg({ ok: false, text: describeError(err) });
    }
  }
  return (
    <form className="card form" onSubmit={save}>
      <h2>Default working hours</h2>
      <p className="muted small">Applied automatically to every Sunday - Thursday without another status.</p>
      <div className="row">
        <label className="field">
          <span>Default start time</span>
          <input type="time" value={start} required onChange={(e) => setStart(e.target.value)} />
        </label>
        <label className="field">
          <span>Default end time</span>
          <input type="time" value={end} required onChange={(e) => setEnd(e.target.value)} />
        </label>
      </div>
      <p className="muted small">Normal working day: {fmt(hoursBetween(start, end))} hours.</p>
      <label className="field">
        <span>Holiday eve end time</span>
        <input type="time" value={eveEnd} required onChange={(e) => setEveEnd(e.target.value)} />
      </label>
      <p className="muted small">
        On a holiday eve (Sunday - Thursday) the automatic day is {start}-{eveEnd} = {fmt(hoursBetween(start, eveEnd))} hours.
      </p>
      {msg && <p className={msg.ok ? 'success' : 'alert'}>{msg.text}</p>}
      <div className="actions">
        <button className="btn btn-primary" type="submit">
          Save working hours
        </button>
      </div>
    </form>
  );
}

function EmployeeRow({ employee }: { employee: Employee | null }) {
  const { saveEmployee } = useApp();
  const [name, setName] = useState(employee?.fullName ?? '');
  const [role, setRole] = useState<Role>(employee?.role ?? 'employee');
  const [active, setActive] = useState(employee?.active ?? true);
  const [pin, setPin] = useState('');
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);
  const isNew = !employee;

  async function save(e: React.FormEvent) {
    e.preventDefault();
    try {
      await saveEmployee({ id: employee?.id ?? null, fullName: name, role, active, pin: pin || null });
      setPin('');
      setMsg({ ok: true, text: isNew ? 'Employee added.' : 'Saved.' });
      if (isNew) {
        setName('');
        setRole('employee');
        setActive(true);
      }
    } catch (err) {
      setMsg({ ok: false, text: describeError(err) });
    }
  }

  return (
    <form className="employee-row" onSubmit={save}>
      <label className="field">
        <span>Name</span>
        <input value={name} onChange={(e) => setName(e.target.value)} required maxLength={100} />
      </label>
      <label className="field">
        <span>Role</span>
        <select value={role} onChange={(e) => setRole(e.target.value as Role)}>
          <option value="employee">Employee</option>
          <option value="admin">Administrator</option>
        </select>
      </label>
      <label className="field">
        <span>{isNew ? 'PIN' : 'New PIN'}</span>
        <input
          type="password"
          inputMode="numeric"
          autoComplete="new-password"
          value={pin}
          placeholder={isNew ? '4-8 digits' : 'Leave empty to keep'}
          onChange={(e) => setPin(e.target.value.replace(/\D/g, ''))}
          maxLength={8}
          required={isNew}
        />
      </label>
      <label className="check">
        <input type="checkbox" checked={active} onChange={(e) => setActive(e.target.checked)} /> Active
      </label>
      <button className="btn btn-primary" type="submit">
        {isNew ? 'Add employee' : 'Save'}
      </button>
      {msg && <p className={`row-msg ${msg.ok ? 'success' : 'alert'}`}>{msg.text}</p>}
    </form>
  );
}

function HolidaysSection() {
  const { data, saveHoliday, deleteHoliday } = useApp();
  const [date, setDate] = useState('');
  const [name, setName] = useState('');
  const [kind, setKind] = useState<Holiday['kind']>('holiday');
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);

  async function add(e: React.FormEvent) {
    e.preventDefault();
    try {
      await saveHoliday({ date, name, kind });
      setMsg({ ok: true, text: 'Saved.' });
      setDate('');
      setName('');
    } catch (err) {
      setMsg({ ok: false, text: describeError(err) });
    }
  }

  async function remove(h: Holiday) {
    if (!window.confirm(`Remove ${h.name} (${formatDate(h.date)})?`)) return;
    try {
      await deleteHoliday(h.date);
    } catch (err) {
      setMsg({ ok: false, text: describeError(err) });
    }
  }

  return (
    <section className="card">
      <h2>Holidays 2026</h2>
      <p className="muted small">
        Holiday: no work and no leave is counted. Holiday eve: work until {data.settings.eveEnd}. Applies to all employees.
      </p>
      <ul className="holiday-list">
        {data.settings.holidays.map((h) => (
          <li key={h.date}>
            <span className="nowrap">
              {formatDate(h.date)} <span className="muted">{DAY_SHORT[dayOfWeek(h.date)]}</span>
            </span>
            <span className="holiday-name">{h.name}</span>
            <span className={`badge ind-${h.kind === 'eve' ? 'partial' : 'holiday'}`}>{h.kind === 'eve' ? 'Eve' : 'Holiday'}</span>
            <button type="button" className="btn btn-small btn-ghost" onClick={() => remove(h)} aria-label={`Remove ${h.name}`}>
              Remove
            </button>
          </li>
        ))}
      </ul>
      <h3>Add or change a holiday</h3>
      <form className="holiday-form" onSubmit={add}>
        <label className="field">
          <span>Date</span>
          <input type="date" min="2026-01-01" max="2026-12-31" value={date} required onChange={(e) => setDate(e.target.value)} />
        </label>
        <label className="field">
          <span>Name</span>
          <input value={name} required maxLength={100} placeholder="e.g. Election Day" onChange={(e) => setName(e.target.value)} />
        </label>
        <label className="field">
          <span>Type</span>
          <select value={kind} onChange={(e) => setKind(e.target.value as Holiday['kind'])}>
            <option value="holiday">Holiday (no work)</option>
            <option value="eve">Holiday eve (short day)</option>
          </select>
        </label>
        <button className="btn btn-primary" type="submit">
          Save holiday
        </button>
        {msg && <p className={`row-msg ${msg.ok ? 'success' : 'alert'}`}>{msg.text}</p>}
      </form>
    </section>
  );
}

export function SettingsPage() {
  const { data, isAdmin } = useApp();
  if (!isAdmin) return <Navigate to="/" replace />;
  return (
    <div className="stack">
      <div className="page-head">
        <h1>Settings</h1>
      </div>
      <WorkHoursForm />
      <HolidaysSection />
      <section className="card">
        <h2>Employees</h2>
        <p className="muted small">PIN codes are stored hashed and cannot be displayed. Enter a new PIN to change it.</p>
        {data.employees.map((e) => (
          <EmployeeRow key={e.id} employee={e} />
        ))}
        <h3>Add employee</h3>
        <EmployeeRow key={`new-${data.employees.length}`} employee={null} />
      </section>
      <section className="card">
        <h2>Audit log</h2>
        <ChangeList employeeId={null} limit={100} />
      </section>
    </div>
  );
}
