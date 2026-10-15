import type { ComputedDay } from '../lib/attendance';
import { fmtHours } from '../lib/attendance';
import { DAY_SHORT, todayIso } from '../lib/dates';
import { dayLabel } from './StatusBadge';

const SHORT: Record<string, string> = {
  present: 'Work',
  vacation: 'Vac',
  sick: 'Sick',
  partial: 'Part',
  manual: 'Man',
  no_work: 'Off',
  weekend: '',
  holiday: 'Hol',
};

/** One month grid, weeks start on Sunday. */
export function MonthCalendar({ days, onSelect }: { days: ComputedDay[]; onSelect(date: string): void }) {
  if (days.length === 0) return null;
  const today = todayIso();
  const blanks = days[0].dayOfWeek;
  return (
    <div className="calendar">
      {DAY_SHORT.map((d, i) => (
        <div key={d} className={`cal-head ${i >= 5 ? 'is-weekend' : ''}`}>
          {d}
        </div>
      ))}
      {Array.from({ length: blanks }, (_, i) => (
        <div key={`b${i}`} className="cal-blank" />
      ))}
      {days.map((day) => {
        const n = Number(day.date.slice(8));
        const hours =
          day.workedHours > 0 && (day.status !== 'present' || day.isException) ? fmtHours(day.workedHours) : '';
        return (
          <button
            key={day.date}
            type="button"
            className={`cal-day ind-${day.indicator} ${day.date === today ? 'is-today' : ''}`}
            onClick={() => onSelect(day.date)}
            aria-label={`${day.date}: ${dayLabel(day)}`}
          >
            <span className="cal-num">{n}</span>
            <span className="cal-status">{dayLabel(day)}</span>
            <span className="cal-short">{day.kind === 'eve' && day.indicator === 'present' ? 'Eve' : SHORT[day.indicator]}</span>
            {hours && <span className="cal-hours">{hours}</span>}
            {day.kind === 'eve' && day.holidayName && <span className="cal-hours cal-eve">{day.holidayName}</span>}
            {day.notes && <span className="cal-note" title={day.notes}>●</span>}
          </button>
        );
      })}
    </div>
  );
}
