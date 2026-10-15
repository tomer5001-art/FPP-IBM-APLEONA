import { useMemo, useState } from 'react';
import { EmployeeSelect, MonthNav } from '../components/Selectors';
import { StatusBadge } from '../components/StatusBadge';
import { describeError, useApp } from '../context/AppContext';
import { useSelectedEmployee } from './CalendarPage';
import { fmt } from '../lib/attendance';
import { DAY_SHORT, defaultMonth, formatDate, MONTH_NAMES } from '../lib/dates';
import { exportCsv, exportPdf } from '../lib/export';
import { buildAnnualReport, buildMonthlyReport, type ReportTable } from '../lib/reports';

function slug(s: string) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
}

function ExportButtons({ tables, name }: { tables: ReportTable[]; name: string }) {
  const [error, setError] = useState('');
  return (
    <div className="export">
      <button type="button" className="btn" onClick={() => exportPdf(tables, `${name}.pdf`).catch((e) => setError(describeError(e)))}>
        Export PDF
      </button>
      <button type="button" className="btn" onClick={() => exportCsv(tables, `${name}.csv`)}>
        Export Excel (CSV)
      </button>
      {error && <span className="alert">{error}</span>}
    </div>
  );
}

function MonthlyReport() {
  const { data, isAdmin, employeeById } = useApp();
  const [employeeId, setEmployeeId] = useSelectedEmployee();
  const [month, setMonth] = useState(defaultMonth());
  const employee = employeeById(employeeId)!;
  const report = useMemo(() => buildMonthlyReport(employee, month, data.records, data.settings), [employee, month, data]);

  return (
    <section className="card">
      <div className="toolbar">
        {isAdmin && <EmployeeSelect employees={data.employees} value={employeeId} onChange={setEmployeeId} />}
        <MonthNav month={month} onChange={setMonth} />
        <ExportButtons tables={[report.table]} name={`attendance-${slug(employee.fullName)}-2026-${String(month).padStart(2, '0')}`} />
      </div>
      <h2>
        {employee.fullName} · {MONTH_NAMES[month - 1]} 2026
      </h2>
      <div className="table-wrap">
        <table className="table report">
          <thead>
            <tr>
              <th>Date</th>
              <th>Day</th>
              <th>Status</th>
              <th>Start</th>
              <th>End</th>
              <th className="num">Worked h</th>
              <th className="num">Vacation h</th>
              <th className="num">Sick h</th>
              <th className="num">Vacation days</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            {report.days.map((d) => (
              <tr key={d.date} className={d.indicator === 'weekend' ? 'row-weekend' : ''}>
                <td className="nowrap">{formatDate(d.date)}</td>
                <td>{DAY_SHORT[d.dayOfWeek]}</td>
                <td>
                  <StatusBadge day={d} />
                </td>
                <td>{d.startTime ?? ''}</td>
                <td>{d.endTime ?? ''}</td>
                <td className="num">{fmt(d.workedHours)}</td>
                <td className="num">{fmt(d.vacationHours)}</td>
                <td className="num">{fmt(d.sickHours)}</td>
                <td className="num">{fmt(d.vacationDays)}</td>
                <td className="notes-cell">{d.notes ?? ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <h3>Monthly totals</h3>
      <dl className="totals">
        {report.table.totals.map(([k, v]) => (
          <div key={k}>
            <dt>{k}</dt>
            <dd>{v}</dd>
          </div>
        ))}
      </dl>
    </section>
  );
}

function AnnualReport() {
  const { data } = useApp();
  const report = useMemo(() => buildAnnualReport(data.employees, data.records, data.settings), [data]);
  const [open, setOpen] = useState<string | null>(data.employees.length === 1 ? data.employees[0].id : null);

  return (
    <section className="card">
      <div className="toolbar">
        <h2>Annual report 2026</h2>
        <ExportButtons tables={[report.summary, ...report.breakdowns]} name="annual-report-2026" />
      </div>
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              {report.summary.headers.map((h, i) => (
                <th key={h} className={i > 0 ? 'num' : ''}>
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {report.perEmployee.map((p, idx) => (
              <tr key={p.employee.id} className="clickable" onClick={() => setOpen(open === p.employee.id ? null : p.employee.id)}>
                {report.summary.rows[idx].map((c, i) => (
                  <td key={i} className={i > 0 ? 'num' : ''}>
                    {c}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="muted small">Select an employee row to see the monthly breakdown.</p>
      {report.breakdowns.map((b, idx) =>
        report.perEmployee[idx].employee.id === open ? (
          <div key={b.title}>
            <h3>{b.title}</h3>
            <div className="table-wrap">
              <table className="table">
                <thead>
                  <tr>
                    {b.headers.map((h, i) => (
                      <th key={h} className={i > 0 ? 'num' : ''}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {b.rows.map((r, ri) => (
                    <tr key={ri} className={ri === b.rows.length - 1 ? 'row-total' : ''}>
                      {r.map((c, i) => (
                        <td key={i} className={i > 0 ? 'num' : ''}>
                          {c}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null,
      )}
    </section>
  );
}

export function ReportsPage() {
  const [tab, setTab] = useState<'monthly' | 'annual'>('monthly');
  return (
    <div className="stack">
      <div className="page-head">
        <h1>Reports</h1>
        <div className="tabs" role="tablist">
          <button type="button" role="tab" aria-selected={tab === 'monthly'} className={tab === 'monthly' ? 'is-active' : ''} onClick={() => setTab('monthly')}>
            Monthly
          </button>
          <button type="button" role="tab" aria-selected={tab === 'annual'} className={tab === 'annual' ? 'is-active' : ''} onClick={() => setTab('annual')}>
            Annual
          </button>
        </div>
      </div>
      {tab === 'monthly' ? <MonthlyReport /> : <AnnualReport />}
    </div>
  );
}
