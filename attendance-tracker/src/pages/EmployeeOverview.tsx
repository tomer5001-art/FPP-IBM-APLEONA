import { useMemo, useState } from 'react';
import { Link } from 'react-router-dom';
import { DayDialog } from '../components/DayDialog';
import { MonthCalendar } from '../components/MonthCalendar';
import { MonthNav } from '../components/Selectors';
import { Legend } from '../components/StatusBadge';
import { StatTile } from '../components/StatTile';
import { useApp } from '../context/AppContext';
import { useEmployeeStats } from '../context/useEmployeeStats';
import { fmt, fmtHours } from '../lib/attendance';
import { defaultMonth, monthOf, MONTH_NAMES } from '../lib/dates';

/** Personal dashboard: used for an employee's own view and by the admin per employee. */
export function EmployeeOverview({ employeeId }: { employeeId: string }) {
  const { isAdmin, employeeById } = useApp();
  const stats = useEmployeeStats(employeeId);
  const [month, setMonth] = useState(defaultMonth());
  const [selected, setSelected] = useState<string | null>(null);
  const days = useMemo(() => stats.year.filter((d) => monthOf(d.date) === month), [stats.year, month]);
  const selectedDay = selected ? days.find((d) => d.date === selected) ?? stats.year.find((d) => d.date === selected) : undefined;
  const q = isAdmin ? `&employee=${employeeId}` : '';

  return (
    <div className="stack">
      <div className="page-head">
        <div>
          <h1>{employeeById(employeeId)?.fullName}</h1>
          <p className="muted">Current month: {MONTH_NAMES[stats.month - 1]} 2026</p>
        </div>
      </div>

      <div className="stats">
        <StatTile label="Vacation days 2026" value={fmt(stats.yearTotals.vacationDays)} hint={`${fmt(stats.yearTotals.vacationHours)} hours`} />
        <StatTile label="Sick leave days 2026" value={fmt(stats.yearTotals.sickDays)} hint={`${fmt(stats.yearTotals.sickHours)} hours`} />
        <StatTile label="Days worked this month" value={stats.monthToDate.daysWorked} hint={`of ${stats.monthTotals.normalWorkingDays} working days`} />
        <StatTile label="Hours this month" value={fmtHours(stats.monthToDate.workedHours)} hint="to date" />
      </div>

      <div className="quick-actions">
        <Link className="btn btn-primary btn-large" to={`/report?type=vacation${q}`}>
          Report Vacation
        </Link>
        <Link className="btn btn-large btn-sick" to={`/report?type=sick${q}`}>
          Report Sick Leave
        </Link>
        <Link className="btn btn-large" to={`/report?type=manual${q}`}>
          Enter Manual Work Hours
        </Link>
      </div>

      <section className="card">
        <div className="card-head">
          <h2>Calendar</h2>
          <MonthNav month={month} onChange={setMonth} />
        </div>
        <MonthCalendar days={days} onSelect={setSelected} />
        <Legend />
      </section>

      {selectedDay && <DayDialog key={selectedDay.date} employeeId={employeeId} day={selectedDay} onClose={() => setSelected(null)} />}
    </div>
  );
}
