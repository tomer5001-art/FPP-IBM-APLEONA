import type { ComputedDay, Indicator } from '../lib/attendance';
import { STATUS_LABELS } from '../lib/types';

export const INDICATOR_LABELS: Record<Indicator, string> = {
  present: 'Present',
  vacation: 'Vacation',
  sick: 'Sick Leave',
  partial: 'Partial Leave',
  manual: 'Manual Work',
  no_work: 'No Work',
  weekend: 'Weekend',
  holiday: 'Holiday',
};

export function dayLabel(day: ComputedDay): string {
  if (day.indicator === 'holiday') return day.holidayName ?? 'Holiday';
  if (day.indicator === 'weekend') return 'Weekend';
  return STATUS_LABELS[day.status];
}

export function StatusBadge({ day }: { day: ComputedDay }) {
  return <span className={`badge ind-${day.indicator}`}>{dayLabel(day)}</span>;
}

export function Legend() {
  const items: Indicator[] = ['present', 'vacation', 'sick', 'partial', 'manual', 'holiday', 'weekend'];
  return (
    <div className="legend" aria-label="Legend">
      {items.map((i) => (
        <span key={i} className="legend-item">
          <span className={`dot ind-${i}`} /> {INDICATOR_LABELS[i]}
        </span>
      ))}
    </div>
  );
}
