-- Holidays and holiday eves.
--  * holiday on Sunday-Thursday: no automatic work, leave cannot be recorded,
--    and it is not counted as a normal working day.
--  * holiday eve on Sunday-Thursday: automatic work from the default start
--    until app_settings.eve_end (14:00 by default).

alter table public.app_settings add column if not exists eve_end time not null default '14:00';

create table if not exists public.holidays (
  holiday_date date primary key check (holiday_date between date '2026-01-01' and date '2026-12-31'),
  name         text not null check (length(trim(name)) between 1 and 100),
  kind         text not null check (kind in ('holiday', 'eve')),
  updated_by   uuid references public.employees (id) on delete set null,
  updated_at   timestamptz not null default now()
);
alter table public.holidays enable row level security;
revoke all on public.holidays from public, anon, authenticated;

alter table public.audit_log drop constraint if exists audit_log_entity_check;
alter table public.audit_log add constraint audit_log_entity_check
  check (entity in ('day', 'settings', 'employee', 'holiday'));

-- Israeli holidays 2026 (5786/5787). Holidays falling on Friday/Saturday are
-- listed too so they show in the calendar.
insert into public.holidays (holiday_date, name, kind) values
  ('2026-04-01', 'Erev Pesach',                  'eve'),
  ('2026-04-02', 'Pesach',                       'holiday'),
  ('2026-04-07', 'Erev Seventh Day of Pesach',   'eve'),
  ('2026-04-08', 'Seventh Day of Pesach',        'holiday'),
  ('2026-04-22', 'Independence Day',             'holiday'),
  ('2026-05-21', 'Erev Shavuot',                 'eve'),
  ('2026-05-22', 'Shavuot',                      'holiday'),
  ('2026-09-11', 'Erev Rosh Hashana',            'eve'),
  ('2026-09-12', 'Rosh Hashana',                 'holiday'),
  ('2026-09-13', 'Rosh Hashana (second day)',    'holiday'),
  ('2026-09-20', 'Erev Yom Kippur',              'eve'),
  ('2026-09-21', 'Yom Kippur',                   'holiday'),
  ('2026-09-25', 'Erev Sukkot',                  'eve'),
  ('2026-09-26', 'Sukkot',                       'holiday'),
  ('2026-10-02', 'Erev Simchat Torah',           'eve'),
  ('2026-10-03', 'Simchat Torah',                'holiday')
on conflict (holiday_date) do nothing;

-- 'workday' | 'eve' | 'holiday' | 'weekend'
create or replace function public._day_kind(p_date date)
returns text
language sql stable
set search_path = public, extensions, pg_temp
as $$
  select case
    when h.kind = 'holiday' then 'holiday'
    when extract(dow from p_date) >= 5 then 'weekend'
    when h.kind = 'eve' then 'eve'
    else 'workday' end
  from (select 1) x
  left join public.holidays h on h.holiday_date = p_date
$$;

create or replace function public.get_app_data(p_token text)
returns jsonb
language plpgsql stable security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me    public.employees;
  v_admin boolean;
begin
  v_me := public._current_employee(p_token);
  v_admin := v_me.role = 'admin';

  return jsonb_build_object(
    'me', jsonb_build_object('id', v_me.id, 'full_name', v_me.full_name, 'role', v_me.role),

    'settings', (select jsonb_build_object(
                    'year', s.year,
                    'default_start', to_char(s.default_start, 'HH24:MI'),
                    'default_end', to_char(s.default_end, 'HH24:MI'),
                    'eve_end', to_char(s.eve_end, 'HH24:MI'))
                   from public.app_settings s where s.id = 1),

    'holidays', coalesce((select jsonb_agg(jsonb_build_object(
                    'date', h.holiday_date, 'name', h.name, 'kind', h.kind)
                    order by h.holiday_date)
                   from public.holidays h), '[]'::jsonb),

    'employees', coalesce((select jsonb_agg(jsonb_build_object(
                    'id', e.id, 'full_name', e.full_name, 'role', e.role, 'active', e.active)
                    order by e.created_at, e.full_name)
                   from public.employees e
                  where v_admin or e.id = v_me.id), '[]'::jsonb),

    'records', coalesce((select jsonb_agg(jsonb_build_object(
                    'employee_id', a.employee_id,
                    'date',        a.work_date,
                    'status',      a.status,
                    'start_time',  to_char(a.start_time, 'HH24:MI'),
                    'end_time',    to_char(a.end_time, 'HH24:MI'),
                    'notes',       a.notes,
                    'leave_type',  l.leave_type,
                    'leave_full_day', l.full_day,
                    'leave_start', to_char(l.start_time, 'HH24:MI'),
                    'leave_end',   to_char(l.end_time, 'HH24:MI'),
                    'leave_hours', l.hours,
                    'updated_at',  a.updated_at)
                    order by a.work_date)
                   from public.attendance_records a
                   left join public.leave_records l
                          on l.employee_id = a.employee_id and l.leave_date = a.work_date
                  where v_admin or a.employee_id = v_me.id), '[]'::jsonb),

    'last_changes', coalesce((select jsonb_agg(to_jsonb(x)) from (
                    select distinct on (al.employee_id)
                           al.employee_id, al.changed_at, al.changed_by_name, al.action, al.record_date
                      from public.audit_log al
                     where al.entity = 'day' and (v_admin or al.employee_id = v_me.id)
                     order by al.employee_id, al.changed_at desc) x), '[]'::jsonb)
  );
end
$$;

-- Same as before, but holiday-aware.
create or replace function public.save_day(
  p_token       text,
  p_employee_id uuid,
  p_date        date,
  p_status      text,
  p_start       text default null,
  p_end         text default null,
  p_leave_start text default null,
  p_leave_end   text default null,
  p_notes       text default null)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me       public.employees;
  v_settings public.app_settings;
  v_kind     text;
  v_workday  boolean;
  v_day_end  time;
  v_start    time := public._parse_time(p_start);
  v_end      time := public._parse_time(p_end);
  v_ls       time := public._parse_time(p_leave_start);
  v_le       time := public._parse_time(p_leave_end);
  v_notes    text := nullif(trim(coalesce(p_notes, '')), '');
  v_full     numeric;
  v_prev     jsonb;
  v_new      jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_can_access(v_me, p_employee_id);

  if p_date is null or p_date < date '2026-01-01' or p_date > date '2026-12-31' then
    raise exception 'date_out_of_range';
  end if;
  if extract(dow from p_date) = 6 and v_me.role <> 'admin' then
    raise exception 'saturday_admin_only';
  end if;
  v_kind := public._day_kind(p_date);
  v_workday := v_kind in ('workday', 'eve');

  if p_status is null or p_status not in
     ('present', 'vacation', 'sick', 'partial_vacation', 'partial_sick', 'manual', 'no_work') then
    raise exception 'invalid_status';
  end if;
  if v_notes is not null and length(v_notes) > 500 then
    raise exception 'notes_too_long';
  end if;

  if (v_start is null) <> (v_end is null) then raise exception 'work_time_incomplete'; end if;
  if v_start is not null and v_end <= v_start then raise exception 'work_time_invalid'; end if;
  if (v_ls is null) <> (v_le is null) then raise exception 'leave_time_incomplete'; end if;
  if v_ls is not null and v_le <= v_ls then raise exception 'leave_time_invalid'; end if;

  if p_status in ('present', 'manual', 'no_work') then
    v_ls := null; v_le := null;
  end if;
  if p_status in ('no_work', 'vacation', 'sick') then
    v_start := null; v_end := null; v_ls := null; v_le := null;
  end if;
  if p_status = 'manual' and v_start is null then
    raise exception 'work_time_required';
  end if;
  if p_status = 'present' and not v_workday and v_start is null then
    raise exception 'work_time_required';
  end if;
  if p_status in ('vacation', 'sick', 'partial_vacation', 'partial_sick') and not v_workday then
    raise exception 'leave_requires_workday';
  end if;
  if p_status in ('partial_vacation', 'partial_sick') then
    if v_ls is null then raise exception 'leave_time_required'; end if;
    if v_start is not null and v_start < v_le and v_ls < v_end then
      raise exception 'time_overlap';
    end if;
  end if;

  select * into v_settings from public.app_settings where id = 1;
  v_day_end := case when v_kind = 'eve' then v_settings.eve_end else v_settings.default_end end;
  v_full := extract(epoch from (v_day_end - v_settings.default_start)) / 3600;
  v_prev := public._day_state(p_employee_id, p_date);

  if v_notes is null and (
       (v_workday and p_status = 'present'
          and (v_start is null or (v_start = v_settings.default_start and v_end = v_day_end)))
       or (not v_workday and p_status = 'no_work')) then
    delete from public.leave_records      where employee_id = p_employee_id and leave_date = p_date;
    delete from public.attendance_records where employee_id = p_employee_id and work_date = p_date;
  else
    insert into public.attendance_records
      (employee_id, work_date, status, start_time, end_time, notes, updated_by, updated_at)
    values (p_employee_id, p_date, p_status, v_start, v_end, v_notes, v_me.id, now())
    on conflict (employee_id, work_date) do update
      set status = excluded.status, start_time = excluded.start_time, end_time = excluded.end_time,
          notes = excluded.notes, updated_by = excluded.updated_by, updated_at = excluded.updated_at;

    delete from public.leave_records where employee_id = p_employee_id and leave_date = p_date;
    if p_status in ('vacation', 'sick', 'partial_vacation', 'partial_sick') then
      insert into public.leave_records
        (employee_id, leave_date, leave_type, full_day, start_time, end_time, hours, updated_by)
      values (
        p_employee_id, p_date,
        case when p_status in ('vacation', 'partial_vacation') then 'vacation' else 'sick' end,
        p_status in ('vacation', 'sick'),
        v_ls, v_le,
        case when p_status in ('vacation', 'sick') then v_full
             else extract(epoch from (v_le - v_ls)) / 3600 end,
        v_me.id);
    end if;
  end if;

  v_new := public._day_state(p_employee_id, p_date);
  if v_prev is distinct from v_new then
    perform public._audit(v_me, p_employee_id, 'day',
                          case when v_new is null then 'reset' else 'update' end,
                          p_date, v_prev, v_new);
  end if;
  return jsonb_build_object('ok', true, 'day', v_new);
end
$$;

-- Settings now include the holiday-eve end time (null keeps the current value).
drop function if exists public.update_settings(text, text, text);
create or replace function public.update_settings(
  p_token text, p_default_start text, p_default_end text, p_eve_end text default null)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me    public.employees;
  v_old   public.app_settings;
  v_start time := public._parse_time(p_default_start);
  v_end   time := public._parse_time(p_default_end);
  v_eve   time := public._parse_time(p_eve_end);
begin
  v_me := public._current_employee(p_token);
  perform public._assert_admin(v_me);
  select * into v_old from public.app_settings where id = 1;
  v_eve := coalesce(v_eve, v_old.eve_end);
  if v_start is null or v_end is null or v_end <= v_start or v_eve <= v_start then
    raise exception 'work_time_invalid';
  end if;

  update public.app_settings
     set default_start = v_start, default_end = v_end, eve_end = v_eve,
         updated_at = now(), updated_by = v_me.id
   where id = 1;

  perform public._audit(v_me, null, 'settings', 'update', null,
    jsonb_build_object('default_start', to_char(v_old.default_start, 'HH24:MI'),
                       'default_end', to_char(v_old.default_end, 'HH24:MI'),
                       'eve_end', to_char(v_old.eve_end, 'HH24:MI')),
    jsonb_build_object('default_start', to_char(v_start, 'HH24:MI'),
                       'default_end', to_char(v_end, 'HH24:MI'),
                       'eve_end', to_char(v_eve, 'HH24:MI')));
  return jsonb_build_object('ok', true);
end
$$;

create or replace function public.save_holiday(p_token text, p_date date, p_name text, p_kind text)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me   public.employees;
  v_name text := trim(coalesce(p_name, ''));
  v_prev jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_admin(v_me);
  if p_date is null or p_date < date '2026-01-01' or p_date > date '2026-12-31' then
    raise exception 'date_out_of_range';
  end if;
  if length(v_name) = 0 or length(v_name) > 100 then raise exception 'invalid_name'; end if;
  if p_kind is null or p_kind not in ('holiday', 'eve') then raise exception 'invalid_kind'; end if;

  select jsonb_build_object('name', name, 'kind', kind) into v_prev
    from public.holidays where holiday_date = p_date;
  insert into public.holidays (holiday_date, name, kind, updated_by, updated_at)
  values (p_date, v_name, p_kind, v_me.id, now())
  on conflict (holiday_date) do update
    set name = excluded.name, kind = excluded.kind, updated_by = excluded.updated_by, updated_at = excluded.updated_at;

  perform public._audit(v_me, null, 'holiday', case when v_prev is null then 'create' else 'update' end,
                        p_date, v_prev, jsonb_build_object('name', v_name, 'kind', p_kind));
  return jsonb_build_object('ok', true);
end
$$;

create or replace function public.delete_holiday(p_token text, p_date date)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me   public.employees;
  v_prev jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_admin(v_me);
  select jsonb_build_object('name', name, 'kind', kind) into v_prev
    from public.holidays where holiday_date = p_date;
  if v_prev is not null then
    delete from public.holidays where holiday_date = p_date;
    perform public._audit(v_me, null, 'holiday', 'delete', p_date, v_prev, null);
  end if;
  return jsonb_build_object('ok', true);
end
$$;

revoke execute on function public._day_kind(date) from public, anon, authenticated;

revoke execute on function
  public.update_settings(text, text, text, text),
  public.save_holiday(text, date, text, text),
  public.delete_holiday(text, date)
from public;

grant execute on function
  public.get_app_data(text),
  public.save_day(text, uuid, date, text, text, text, text, text, text),
  public.update_settings(text, text, text, text),
  public.save_holiday(text, date, text, text),
  public.delete_holiday(text, date)
to anon, authenticated;
