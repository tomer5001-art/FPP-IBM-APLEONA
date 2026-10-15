import { useState } from 'react';
import { Navigate } from 'react-router-dom';
import { ChangeList } from '../components/ChangeList';
import { describeError, useApp } from '../context/AppContext';
import { fmt, hoursBetween } from '../lib/attendance';
import type { Employee, Role } from '../lib/types';

function WorkHoursForm() {
  const { data, updateSettings } = useApp();
  const [start, setStart] = useState(data.settings.defaultStart);
  const [end, setEnd] = useState(data.settings.defaultEnd);
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null);

  async function save(e: React.FormEvent) {
    e.preventDefault();
    try {
      await updateSettings(start, end);
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

export function SettingsPage() {
  const { data, isAdmin } = useApp();
  if (!isAdmin) return <Navigate to="/" replace />;
  return (
    <div className="stack">
      <div className="page-head">
        <h1>Settings</h1>
      </div>
      <WorkHoursForm />
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
