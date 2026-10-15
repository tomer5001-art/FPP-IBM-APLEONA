import type { AppData, AuditEntry, Employee, SaveDayInput, SaveEmployeeInput } from '../lib/types';

export class ApiError extends Error {
  constructor(public code: string) {
    super(code);
  }
}

export interface LoginResult {
  token: string;
  employee: Employee;
}

/** Backend contract. Implemented by Supabase (production) and a local demo store. */
export interface Api {
  readonly mode: 'supabase' | 'demo';
  listLoginEmployees(): Promise<{ id: string; fullName: string }[]>;
  login(employeeId: string, pin: string): Promise<LoginResult>;
  logout(token: string): Promise<void>;
  getAppData(token: string): Promise<AppData>;
  saveDay(token: string, input: SaveDayInput): Promise<void>;
  resetDay(token: string, employeeId: string, date: string): Promise<void>;
  updateSettings(token: string, defaultStart: string, defaultEnd: string): Promise<void>;
  saveEmployee(token: string, input: SaveEmployeeInput): Promise<string>;
  getAuditLog(token: string, employeeId: string | null, limit?: number): Promise<AuditEntry[]>;
}
