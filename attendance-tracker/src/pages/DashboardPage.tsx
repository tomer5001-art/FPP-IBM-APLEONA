import { Link } from 'react-router-dom';
import { useApp } from '../context/AppContext';
import { useEmployeeStats } from '../context/useEmployeeStats';
import { fmt } from '../lib/attendance';
import { defaultMonth, formatDate, formatDateTime, MONTH_NAMES } from '../lib/dates';
import type { Employee } from '../lib/types';
import { EmployeeOverview } from './EmployeeOverview';

function EmployeeCard({ employee }: { employee: Employee }) {
  const s = useEmployeeStats(employee.id);
  return (
    <Link to={`/employee/${employee.id}`} className="card employee-card">
      <h3>
        {employee.fullName}
        {!employee.active && <span className="tag tag-muted">Inactive</span>}
      </h3>
      <dl>
        <dt>Working days this month</dt>
        <dd>
          {s.monthToDate.daysWorked} <span className="muted">/ {s.monthTotals.normalWorkingDays}</span>
        </dd>
        <dt>Vacation days 2026</dt>
        <dd>{fmt(s.yearTotals.vacationDays)}</dd>
        <dt>Sick leave days 2026</dt>
        <dd>{fmt(s.yearTotals.sickDays)}</dd>
        <dt>Last attendance change</dt>
        <dd className="small">
          {s.lastChange
            ? `${formatDateTime(s.lastChange.changedAt)} by ${s.lastChange.changedByName}${s.lastChange.recordDate ? ` (day ${formatDate(s.lastChange.recordDate)})` : ''}`
            : 'None'}
        </dd>
      </dl>
      <span className="link">View full record →</span>
    </Link>
  );
}

function VacationRow({ employee }: { employee: Employee }) {
  const s = useEmployeeStats(employee.id);
  return (
    <li>
      <span>{employee.fullName}</span>
      <strong>{fmt(s.yearTotals.vacationDays)} days</strong>
    </li>
  );
}

function AdminDashboard() {
  const { data } = useApp();
  const month = MONTH_NAMES[defaultMonth() - 1];
  return (
    <div className="stack">
      <div className="page-head">
        <div>
          <h1>Dashboard</h1>
          <p className="muted">{month} 2026 · working days to date / scheduled</p>
        </div>
      </div>
      <section className="card">
        <h2>Vacation summary 2026</h2>
        <ul className="summary-list">
          {data.employees.filter((e) => e.active).map((e) => (
            <VacationRow key={e.id} employee={e} />
          ))}
        </ul>
      </section>
      <div className="cards">
        {data.employees.map((e) => (
          <EmployeeCard key={e.id} employee={e} />
        ))}
      </div>
    </div>
  );
}

export function DashboardPage() {
  const { data, isAdmin } = useApp();
  return isAdmin ? <AdminDashboard /> : <EmployeeOverview employeeId={data.me.id} />;
}
