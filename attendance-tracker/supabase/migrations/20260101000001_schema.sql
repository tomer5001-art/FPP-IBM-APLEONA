-- Attendance & Vacation Tracker 2026 - schema
-- Only exceptions are stored. A normal Sunday-Thursday day without a row is
-- treated as Present (default start - default end). Friday/Saturday without a
-- row are No Work.

create schema if not exists extensions;
create extension if not exists pgcrypto with schema extensions;

-- Employees -------------------------------------------------------------
create table if not exists public.employees (
  id              uuid primary key default gen_random_uuid(),
  full_name       text not null check (length(trim(full_name)) between 1 and 100),
  role            text not null default 'employee' check (role in ('admin', 'employee')),
  pin_hash        text not null,                       -- bcrypt hash, never the PIN itself
  active          boolean not null default true,
  failed_attempts integer not null default 0,
  locked_until    timestamptz,
  created_at      timestamptz not null default clock_timestamp()
);
create unique index if not exists employees_full_name_key on public.employees (lower(full_name));

-- Application settings (single row) --------------------------------------
create table if not exists public.app_settings (
  id            integer primary key default 1 check (id = 1),
  year          integer not null default 2026,
  default_start time not null default '08:00',
  default_end   time not null default '16:30',
  updated_at    timestamptz not null default now(),
  updated_by    uuid references public.employees (id) on delete set null,
  check (default_end > default_start)
);

-- Attendance exceptions (one row per employee per day, only when not default)
create table if not exists public.attendance_records (
  id          uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.employees (id) on delete cascade,
  work_date   date not null check (work_date between date '2026-01-01' and date '2026-12-31'),
  status      text not null check (status in
                ('present', 'vacation', 'sick', 'partial_vacation', 'partial_sick', 'manual', 'no_work')),
  start_time  time,
  end_time    time,
  notes       text check (notes is null or length(notes) <= 500),
  updated_by  uuid references public.employees (id) on delete set null,
  updated_at  timestamptz not null default now(),
  unique (employee_id, work_date),
  check ((start_time is null) = (end_time is null)),
  check (end_time is null or end_time > start_time)
);

-- Leave records (vacation / sick, full or partial day) --------------------
create table if not exists public.leave_records (
  id          uuid primary key default gen_random_uuid(),
  employee_id uuid not null references public.employees (id) on delete cascade,
  leave_date  date not null check (leave_date between date '2026-01-01' and date '2026-12-31'),
  leave_type  text not null check (leave_type in ('vacation', 'sick')),
  full_day    boolean not null,
  start_time  time,
  end_time    time,
  hours       numeric(5, 2) not null check (hours > 0),
  updated_by  uuid references public.employees (id) on delete set null,
  updated_at  timestamptz not null default now(),
  unique (employee_id, leave_date),
  check (full_day or (start_time is not null and end_time is not null and end_time > start_time))
);

-- Audit log ----------------------------------------------------------------
create table if not exists public.audit_log (
  id              bigint generated always as identity primary key,
  changed_at      timestamptz not null default now(),
  changed_by      uuid references public.employees (id) on delete set null,
  changed_by_name text not null,
  employee_id     uuid references public.employees (id) on delete set null,
  entity          text not null check (entity in ('day', 'settings', 'employee')),
  action          text not null,
  record_date     date,
  previous_value  jsonb,
  new_value       jsonb
);
create index if not exists audit_log_employee_idx on public.audit_log (employee_id, changed_at desc);

-- Login sessions (token is stored hashed) ----------------------------------
create table if not exists public.sessions (
  token_hash  text primary key,
  employee_id uuid not null references public.employees (id) on delete cascade,
  created_at  timestamptz not null default now(),
  expires_at  timestamptz not null
);
create index if not exists sessions_employee_idx on public.sessions (employee_id);

-- Lock the tables down. The browser never reads tables directly: every
-- request goes through the SECURITY DEFINER functions in the next migration,
-- which validate the session token and the caller's role on the server.
alter table public.employees          enable row level security;
alter table public.app_settings       enable row level security;
alter table public.attendance_records enable row level security;
alter table public.leave_records      enable row level security;
alter table public.audit_log          enable row level security;
alter table public.sessions           enable row level security;

revoke all on public.employees, public.app_settings, public.attendance_records,
  public.leave_records, public.audit_log, public.sessions from public, anon, authenticated;
