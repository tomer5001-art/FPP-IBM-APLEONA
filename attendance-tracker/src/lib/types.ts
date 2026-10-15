export type Role = 'admin' | 'employee';

export type DayStatus =
  | 'present'
  | 'vacation'
  | 'sick'
  | 'partial_vacation'
  | 'partial_sick'
  | 'manual'
  | 'no_work';

export interface Employee {
  id: string;
  fullName: string;
  role: Role;
  active: boolean;
}

export interface Holiday {
  date: string; // "YYYY-MM-DD"
  name: string;
  kind: 'holiday' | 'eve';
}

export interface Settings {
  year: number;
  defaultStart: string; // "HH:MM"
  defaultEnd: string;
  eveEnd: string; // end of the working day on a holiday eve
  holidays: Holiday[];
}

/** A stored exception for one employee-day. Days without one use the default rules. */
export interface DayRecord {
  employeeId: string;
  date: string; // "YYYY-MM-DD"
  status: DayStatus;
  startTime: string | null;
  endTime: string | null;
  notes: string | null;
  leaveType: 'vacation' | 'sick' | null;
  leaveFullDay: boolean | null;
  leaveStart: string | null;
  leaveEnd: string | null;
  leaveHours: number | null;
  updatedAt: string | null;
}

export interface SaveDayInput {
  employeeId: string;
  date: string;
  status: DayStatus;
  startTime?: string | null;
  endTime?: string | null;
  leaveStart?: string | null;
  leaveEnd?: string | null;
  notes?: string | null;
}

export interface LastChange {
  employeeId: string;
  changedAt: string;
  changedByName: string;
  action: string;
  recordDate: string | null;
}

export interface AuditEntry {
  id: number | string;
  changedAt: string;
  changedByName: string;
  employeeId: string | null;
  employeeName: string | null;
  entity: 'day' | 'settings' | 'employee' | 'holiday';
  action: string;
  recordDate: string | null;
  previousValue: unknown;
  newValue: unknown;
}

export interface AppData {
  me: Employee;
  settings: Settings;
  employees: Employee[];
  records: DayRecord[];
  lastChanges: LastChange[];
}

export interface SaveEmployeeInput {
  id: string | null;
  fullName: string;
  role: Role;
  active: boolean;
  pin?: string | null;
}

export const STATUS_LABELS: Record<DayStatus, string> = {
  present: 'Present',
  vacation: 'Vacation',
  sick: 'Sick Leave',
  partial_vacation: 'Partial Vacation',
  partial_sick: 'Partial Sick Leave',
  manual: 'Manual Work',
  no_work: 'No Work',
};
