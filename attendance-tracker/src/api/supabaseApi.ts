import { createClient, type SupabaseClient } from '@supabase/supabase-js';
import type { Api, LoginResult } from './api';
import { ApiError } from './api';
import { mapAppData, mapAudit, mapEmployee } from './mappers';

/**
 * Talks to Supabase only through RPC functions. Tables are not readable with
 * the public key; every function validates the session token server side.
 */
export function createSupabaseApi(url: string, key: string): Api {
  const client: SupabaseClient = createClient(url, key, {
    auth: { persistSession: false, autoRefreshToken: false },
  });

  async function rpc<T = unknown>(fn: string, args: Record<string, unknown> = {}): Promise<T> {
    const { data, error } = await client.rpc(fn, args);
    if (error) {
      // Our functions raise short codes like "forbidden"; anything else is a network/db error.
      const code = /^[a-z_]+$/.test(error.message) ? error.message : `Server error: ${error.message}`;
      throw new ApiError(code);
    }
    return data as T;
  }

  return {
    mode: 'supabase',

    async listLoginEmployees() {
      const rows = await rpc<{ id: string; full_name: string }[]>('list_login_employees');
      return rows.map((r) => ({ id: r.id, fullName: r.full_name }));
    },

    async login(employeeId, pin): Promise<LoginResult> {
      const res = await rpc<{ error?: string; token?: string; employee?: unknown }>('login', {
        p_employee_id: employeeId,
        p_pin: pin,
      });
      if (res.error || !res.token) throw new ApiError(res.error ?? 'invalid_credentials');
      return { token: res.token, employee: mapEmployee(res.employee) };
    },

    async logout(token) {
      await rpc('logout', { p_token: token });
    },

    async getAppData(token) {
      return mapAppData(await rpc('get_app_data', { p_token: token }));
    },

    async saveDay(token, i) {
      await rpc('save_day', {
        p_token: token,
        p_employee_id: i.employeeId,
        p_date: i.date,
        p_status: i.status,
        p_start: i.startTime || null,
        p_end: i.endTime || null,
        p_leave_start: i.leaveStart || null,
        p_leave_end: i.leaveEnd || null,
        p_notes: i.notes || null,
      });
    },

    async resetDay(token, employeeId, date) {
      await rpc('reset_day', { p_token: token, p_employee_id: employeeId, p_date: date });
    },

    async updateSettings(token, defaultStart, defaultEnd, eveEnd) {
      await rpc('update_settings', { p_token: token, p_default_start: defaultStart, p_default_end: defaultEnd, p_eve_end: eveEnd });
    },

    async saveHoliday(token, h) {
      await rpc('save_holiday', { p_token: token, p_date: h.date, p_name: h.name, p_kind: h.kind });
    },

    async deleteHoliday(token, date) {
      await rpc('delete_holiday', { p_token: token, p_date: date });
    },

    async saveEmployee(token, i) {
      const res = await rpc<{ id: string }>('save_employee', {
        p_token: token,
        p_id: i.id,
        p_full_name: i.fullName,
        p_role: i.role,
        p_active: i.active,
        p_pin: i.pin || null,
      });
      return res.id;
    },

    async getAuditLog(token, employeeId, limit = 50) {
      const rows = await rpc<unknown[]>('get_audit_log', {
        p_token: token,
        p_employee_id: employeeId,
        p_limit: limit,
      });
      return rows.map(mapAudit);
    },
  };
}
