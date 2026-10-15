import type { Api } from './api';
import { createDemoApi } from './demoApi';
import { createSupabaseApi } from './supabaseApi';

const url = import.meta.env.VITE_SUPABASE_URL as string | undefined;
const key = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined;

/** Supabase when configured, otherwise the in-browser demo backend. */
export const api: Api = url && key ? createSupabaseApi(url, key) : createDemoApi();

export { ApiError } from './api';
export type { Api } from './api';
