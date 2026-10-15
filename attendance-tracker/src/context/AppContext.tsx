import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';
import { api, ApiError } from '../api';
import { errorMessage } from '../lib/attendance';
import type { AppData, Employee, SaveDayInput, SaveEmployeeInput } from '../lib/types';

const TOKEN_KEY = 'attendance-2026-token';

interface AppContextValue {
  data: AppData;
  isAdmin: boolean;
  employeeById(id: string): Employee | undefined;
  refresh(): Promise<void>;
  saveDay(input: SaveDayInput): Promise<void>;
  saveDays(inputs: SaveDayInput[]): Promise<void>;
  resetDay(employeeId: string, date: string): Promise<void>;
  updateSettings(start: string, end: string): Promise<void>;
  saveEmployee(input: SaveEmployeeInput): Promise<void>;
  token: string;
  logout(): void;
}

const Ctx = createContext<AppContextValue | null>(null);

export function useApp(): AppContextValue {
  const v = useContext(Ctx);
  if (!v) throw new Error('useApp must be used inside AppProvider');
  return v;
}

function readToken(): string | null {
  try {
    return localStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}

function writeToken(token: string | null) {
  try {
    if (token) localStorage.setItem(TOKEN_KEY, token);
    else localStorage.removeItem(TOKEN_KEY);
  } catch {
    /* storage unavailable - session lasts for this tab only */
  }
}

/** Convert any thrown error into a user-facing message. */
export function describeError(e: unknown): string {
  if (e instanceof ApiError) return errorMessage(e.code);
  if (e instanceof Error) return e.message;
  return String(e);
}

interface ProviderProps {
  children: (state: { status: 'loading' | 'signed-out' | 'ready'; login(token: string): void }) => ReactNode;
  signedIn: ReactNode;
}

/**
 * Holds the session and all data visible to the signed-in employee.
 * Data is small (only exceptions for one year), so it is loaded in one call
 * and reloaded after every change.
 */
export function AppProvider({ children, signedIn }: ProviderProps) {
  const [token, setToken] = useState<string | null>(readToken);
  const [data, setData] = useState<AppData | null>(null);
  const [loading, setLoading] = useState<boolean>(!!token);

  const logout = useCallback(() => {
    if (token) api.logout(token).catch(() => undefined);
    writeToken(null);
    setToken(null);
    setData(null);
  }, [token]);

  const load = useCallback(async (t: string) => {
    try {
      setData(await api.getAppData(t));
    } catch (e) {
      if (e instanceof ApiError && e.code === 'not_authenticated') {
        writeToken(null);
        setToken(null);
        setData(null);
      }
      throw e;
    }
  }, []);

  useEffect(() => {
    if (!token) return;
    setLoading(true);
    load(token)
      .catch(() => undefined)
      .finally(() => setLoading(false));
  }, [token, load]);

  const login = useCallback((t: string) => {
    writeToken(t);
    setToken(t);
  }, []);

  const value = useMemo<AppContextValue | null>(() => {
    if (!token || !data) return null;
    const run = async (fn: () => Promise<unknown>) => {
      try {
        await fn();
      } finally {
        await load(token).catch(() => undefined);
      }
    };
    return {
      data,
      token,
      isAdmin: data.me.role === 'admin',
      employeeById: (id) => data.employees.find((e) => e.id === id),
      refresh: () => load(token),
      saveDay: (input) => run(() => api.saveDay(token, input)),
      saveDays: (inputs) =>
        run(async () => {
          for (const input of inputs) await api.saveDay(token, input);
        }),
      resetDay: (employeeId, date) => run(() => api.resetDay(token, employeeId, date)),
      updateSettings: (s, e) => run(() => api.updateSettings(token, s, e)),
      saveEmployee: (input) => run(() => api.saveEmployee(token, input)),
      logout,
    };
  }, [token, data, load, logout]);

  if (value) return <Ctx.Provider value={value}>{signedIn}</Ctx.Provider>;
  return <>{children({ status: loading ? 'loading' : 'signed-out', login })}</>;
}
