import { useMemo, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import { DayDialog } from '../components/DayDialog';
import { MonthCalendar } from '../components/MonthCalendar';
import { EmployeeSelect, MonthNav } from '../components/Selectors';
import { Legend } from '../components/StatusBadge';
import { useApp } from '../context/AppContext';
import { computeMonth, fmt, fmtHours, summarize } from '../lib/attendance';
import { defaultMonth } from '../lib/dates';

/** Returns the employee to show: admin may choose via ?employee=, others always see themselves. */
export function useSelectedEmployee(): [string, (id: string) => void] {
  const { data, isAdmin } = useApp();
  const [params, setParams] = useSearchParams();
  const requested = params.get('employee');
  const id = isAdmin && requested && data.employees.some((e) => e.id === requested) ? requested : data.me.id;
  const set = (next: string) => {
    const p = new URLSearchParams(params);
    p.set('employee', next);
    setParams(p, { replace: true });
  };
  return [id, set];
}

export function CalendarPage() {
  const { data, isAdmin } = useApp();
  const [employeeId, setEmployeeId] = useSelectedEmployee();
  const [month, setMonth] = useState(defaultMonth());
  const [selected, setSelected] = useState<string | null>(null);
  const days = useMemo(() => computeMonth(month, data.records, employeeId, data.settings), [month, data, employeeId]);
  const totals = summarize(days);
  const selectedDay = days.find((d) => d.date === selected);

  return (
    <div className="stack">
      <div className="page-head">
        <h1>Calendar</h1>
      </div>
      <section className="card">
        <div className="toolbar">
          {isAdmin && <EmployeeSelect employees={data.employees} value={employeeId} onChange={setEmployeeId} />}
          <MonthNav month={month} onChange={setMonth} />
        </div>
        <MonthCalendar days={days} onSelect={setSelected} />
        <Legend />
        <p className="muted small">
          Month: {totals.daysWorked} days worked · {fmtHours(totals.workedHours)} hours · {fmt(totals.vacationDays)} vacation days ·{' '}
          {fmt(totals.sickDays)} sick days. Tap a day to open the daily record.
        </p>
      </section>
      {selectedDay && <DayDialog key={selectedDay.date} employeeId={employeeId} day={selectedDay} onClose={() => setSelected(null)} />}
    </div>
  );
}
