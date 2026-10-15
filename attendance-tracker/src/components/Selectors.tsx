import { MONTH_NAMES } from '../lib/dates';
import type { Employee } from '../lib/types';

export function EmployeeSelect({ employees, value, onChange, label = 'Employee' }: {
  employees: Employee[];
  value: string;
  onChange(id: string): void;
  label?: string;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <select value={value} onChange={(e) => onChange(e.target.value)}>
        {employees.map((e) => (
          <option key={e.id} value={e.id}>
            {e.fullName}
            {e.active ? '' : ' (inactive)'}
          </option>
        ))}
      </select>
    </label>
  );
}

export function MonthNav({ month, onChange }: { month: number; onChange(m: number): void }) {
  return (
    <div className="month-nav">
      <button type="button" className="btn btn-icon" onClick={() => onChange(month - 1)} disabled={month <= 1} aria-label="Previous month">
        ‹
      </button>
      <select value={month} onChange={(e) => onChange(Number(e.target.value))} aria-label="Month">
        {MONTH_NAMES.map((name, i) => (
          <option key={name} value={i + 1}>
            {name} 2026
          </option>
        ))}
      </select>
      <button type="button" className="btn btn-icon" onClick={() => onChange(month + 1)} disabled={month >= 12} aria-label="Next month">
        ›
      </button>
    </div>
  );
}
