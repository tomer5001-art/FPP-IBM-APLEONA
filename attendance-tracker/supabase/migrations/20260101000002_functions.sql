-- Attendance & Vacation Tracker 2026 - server-side API (RPC functions)
--
-- All permission checks happen here, on the server. The client sends a
-- session token obtained from login(); every function resolves the token to
-- an employee and checks that employee's role before touching any data.
-- Errors are raised with short codes (e.g. 'forbidden') that the client maps
-- to friendly messages.

-- Internal helpers ----------------------------------------------------------

create or replace function public._hash_token(p_token text)
returns text
language sql immutable
set search_path = public, extensions, pg_temp
as $$
  select encode(extensions.digest(p_token, 'sha256'), 'hex')
$$;

create or replace function public._current_employee(p_token text)
returns public.employees
language plpgsql stable
set search_path = public, extensions, pg_temp
as $$
declare
  v_emp public.employees;
begin
  if p_token is null or length(p_token) < 32 then
    raise exception 'not_authenticated';
  end if;
  select e.* into v_emp
    from public.sessions s
    join public.employees e on e.id = s.employee_id
   where s.token_hash = public._hash_token(p_token)
     and s.expires_at > now()
     and e.active;
  if not found then
    raise exception 'not_authenticated';
  end if;
  return v_emp;
end
$$;

create or replace function public._assert_admin(p_me public.employees)
returns void
language plpgsql immutable
as $$
begin
  if p_me.role <> 'admin' then
    raise exception 'forbidden';
  end if;
end
$$;

-- Admin may access everyone; an employee may access only their own records.
create or replace function public._assert_can_access(p_me public.employees, p_employee_id uuid)
returns void
language plpgsql stable
set search_path = public, extensions, pg_temp
as $$
begin
  if p_employee_id is null then
    raise exception 'employee_required';
  end if;
  if p_me.role <> 'admin' and p_me.id <> p_employee_id then
    raise exception 'forbidden';
  end if;
  if not exists (select 1 from public.employees where id = p_employee_id) then
    raise exception 'employee_not_found';
  end if;
end
$$;

-- Current stored state of one employee-day (null = automatic default).
create or replace function public._day_state(p_employee_id uuid, p_date date)
returns jsonb
language sql stable
set search_path = public, extensions, pg_temp
as $$
  select case when a.id is null and l.id is null then null else jsonb_build_object(
    'status',     a.status,
    'start_time', to_char(a.start_time, 'HH24:MI'),
    'end_time',   to_char(a.end_time, 'HH24:MI'),
    'notes',      a.notes,
    'leave', case when l.id is null then null else jsonb_build_object(
      'type',       l.leave_type,
      'full_day',   l.full_day,
      'start_time', to_char(l.start_time, 'HH24:MI'),
      'end_time',   to_char(l.end_time, 'HH24:MI'),
      'hours',      l.hours) end
  ) end
  from (select 1) x
  left join public.attendance_records a on a.employee_id = p_employee_id and a.work_date = p_date
  left join public.leave_records l      on l.employee_id = p_employee_id and l.leave_date = p_date
$$;

create or replace function public._audit(
  p_me public.employees, p_employee_id uuid, p_entity text, p_action text,
  p_date date, p_prev jsonb, p_new jsonb)
returns void
language sql
set search_path = public, extensions, pg_temp
as $$
  insert into public.audit_log
    (changed_by, changed_by_name, employee_id, entity, action, record_date, previous_value, new_value)
  values (p_me.id, p_me.full_name, p_employee_id, p_entity, p_action, p_date, p_prev, p_new)
$$;

create or replace function public._parse_time(p_value text)
returns time
language plpgsql immutable
as $$
begin
  if p_value is null or trim(p_value) = '' then
    return null;
  end if;
  if trim(p_value) !~ '^([01][0-9]|2[0-3]):[0-5][0-9]$' then
    raise exception 'invalid_time';
  end if;
  return trim(p_value)::time;
end
$$;

-- Public API ------------------------------------------------------------------

-- Names for the login screen (active employees only, nothing sensitive).
create or replace function public.list_login_employees()
returns jsonb
language sql stable security definer
set search_path = public, extensions, pg_temp
as $$
  select coalesce(jsonb_agg(jsonb_build_object('id', id, 'full_name', full_name)
                            order by created_at, full_name), '[]'::jsonb)
    from public.employees where active
$$;

-- Verify PIN, return a session token. Returns {error} instead of raising so
-- the failed-attempt counter is not rolled back. 5 wrong PINs lock for 5 min.
create or replace function public.login(p_employee_id uuid, p_pin text)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_emp      public.employees;
  v_attempts integer;
  v_token    text;
begin
  select * into v_emp from public.employees where id = p_employee_id and active for update;
  if not found then
    return jsonb_build_object('error', 'invalid_credentials');
  end if;
  if v_emp.locked_until is not null and v_emp.locked_until > now() then
    return jsonb_build_object('error', 'locked', 'locked_until', v_emp.locked_until);
  end if;

  if p_pin is null or v_emp.pin_hash <> extensions.crypt(p_pin, v_emp.pin_hash) then
    v_attempts := v_emp.failed_attempts + 1;
    if v_attempts >= 5 then
      update public.employees
         set failed_attempts = 0, locked_until = now() + interval '5 minutes'
       where id = v_emp.id;
      return jsonb_build_object('error', 'locked');
    end if;
    update public.employees set failed_attempts = v_attempts where id = v_emp.id;
    return jsonb_build_object('error', 'invalid_credentials');
  end if;

  update public.employees set failed_attempts = 0, locked_until = null where id = v_emp.id;
  delete from public.sessions where expires_at < now();

  v_token := encode(extensions.gen_random_bytes(32), 'hex');
  insert into public.sessions (token_hash, employee_id, expires_at)
  values (public._hash_token(v_token), v_emp.id, now() + interval '30 days');

  return jsonb_build_object(
    'token', v_token,
    'employee', jsonb_build_object('id', v_emp.id, 'full_name', v_emp.full_name, 'role', v_emp.role));
end
$$;

create or replace function public.logout(p_token text)
returns jsonb
language sql volatile security definer
set search_path = public, extensions, pg_temp
as $$
  delete from public.sessions where token_hash = public._hash_token(p_token);
  select jsonb_build_object('ok', true);
$$;

-- Everything the UI needs, filtered by role: admin sees all employees,
-- an employee sees only themself.
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
                    'default_end', to_char(s.default_end, 'HH24:MI'))
                   from public.app_settings s where s.id = 1),

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

-- Create / replace the record for one employee-day.
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
  v_dow      integer;
  v_workday  boolean;
  v_start    time := public._parse_time(p_start);
  v_end      time := public._parse_time(p_end);
  v_ls       time := public._parse_time(p_leave_start);
  v_le       time := public._parse_time(p_leave_end);
  v_notes    text := nullif(trim(coalesce(p_notes, '')), '');
  v_normal   numeric;
  v_prev     jsonb;
  v_new      jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_can_access(v_me, p_employee_id);

  if p_date is null or p_date < date '2026-01-01' or p_date > date '2026-12-31' then
    raise exception 'date_out_of_range';
  end if;
  v_dow := extract(dow from p_date)::integer;          -- 0 = Sunday ... 6 = Saturday
  v_workday := v_dow between 0 and 4;                   -- Sunday - Thursday
  if v_dow = 6 and v_me.role <> 'admin' then
    raise exception 'saturday_admin_only';
  end if;
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
  v_normal := extract(epoch from (v_settings.default_end - v_settings.default_start)) / 3600;
  v_prev := public._day_state(p_employee_id, p_date);

  if v_notes is null and (
       (v_workday and p_status = 'present'
          and (v_start is null or (v_start = v_settings.default_start and v_end = v_settings.default_end)))
       or (not v_workday and p_status = 'no_work')) then
    -- Same as the automatic default: store nothing.
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
        case when p_status in ('vacation', 'sick') then v_normal
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

-- Remove the exception so the day goes back to the automatic default.
create or replace function public.reset_day(p_token text, p_employee_id uuid, p_date date)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me   public.employees;
  v_prev jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_can_access(v_me, p_employee_id);
  if p_date is null or p_date < date '2026-01-01' or p_date > date '2026-12-31' then
    raise exception 'date_out_of_range';
  end if;
  if extract(dow from p_date) = 6 and v_me.role <> 'admin' then
    raise exception 'saturday_admin_only';
  end if;

  v_prev := public._day_state(p_employee_id, p_date);
  delete from public.leave_records      where employee_id = p_employee_id and leave_date = p_date;
  delete from public.attendance_records where employee_id = p_employee_id and work_date = p_date;
  if v_prev is not null then
    perform public._audit(v_me, p_employee_id, 'day', 'reset', p_date, v_prev, null);
  end if;
  return jsonb_build_object('ok', true);
end
$$;

create or replace function public.update_settings(p_token text, p_default_start text, p_default_end text)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me    public.employees;
  v_start time := public._parse_time(p_default_start);
  v_end   time := public._parse_time(p_default_end);
  v_prev  jsonb;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_admin(v_me);
  if v_start is null or v_end is null or v_end <= v_start then
    raise exception 'work_time_invalid';
  end if;

  select jsonb_build_object('default_start', to_char(default_start, 'HH24:MI'),
                            'default_end', to_char(default_end, 'HH24:MI'))
    into v_prev from public.app_settings where id = 1;

  update public.app_settings
     set default_start = v_start, default_end = v_end, updated_at = now(), updated_by = v_me.id
   where id = 1;

  perform public._audit(v_me, null, 'settings', 'update', null, v_prev,
    jsonb_build_object('default_start', to_char(v_start, 'HH24:MI'), 'default_end', to_char(v_end, 'HH24:MI')));
  return jsonb_build_object('ok', true);
end
$$;

-- Add (p_id null) or update an employee. p_pin null keeps the current PIN.
create or replace function public.save_employee(
  p_token     text,
  p_id        uuid,
  p_full_name text,
  p_role      text,
  p_active    boolean,
  p_pin       text default null)
returns jsonb
language plpgsql volatile security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me   public.employees;
  v_old  public.employees;
  v_name text := trim(coalesce(p_full_name, ''));
  v_pin  text := nullif(trim(coalesce(p_pin, '')), '');
  v_id   uuid;
begin
  v_me := public._current_employee(p_token);
  perform public._assert_admin(v_me);

  if length(v_name) = 0 or length(v_name) > 100 then raise exception 'invalid_name'; end if;
  if p_role is null or p_role not in ('admin', 'employee') then raise exception 'invalid_role'; end if;
  if v_pin is not null and v_pin !~ '^[0-9]{4,8}$' then raise exception 'invalid_pin_format'; end if;

  begin
    if p_id is null then
      if v_pin is null then raise exception 'pin_required'; end if;
      insert into public.employees (full_name, role, active, pin_hash)
      values (v_name, p_role, coalesce(p_active, true), extensions.crypt(v_pin, extensions.gen_salt('bf', 10)))
      returning id into v_id;
      perform public._audit(v_me, v_id, 'employee', 'create', null, null,
        jsonb_build_object('full_name', v_name, 'role', p_role, 'active', coalesce(p_active, true)));
    else
      select * into v_old from public.employees where id = p_id for update;
      if not found then raise exception 'employee_not_found'; end if;
      update public.employees
         set full_name = v_name,
             role = p_role,
             active = coalesce(p_active, active),
             pin_hash = case when v_pin is null then pin_hash
                             else extensions.crypt(v_pin, extensions.gen_salt('bf', 10)) end,
             failed_attempts = case when v_pin is null then failed_attempts else 0 end,
             locked_until = case when v_pin is null then locked_until else null end
       where id = p_id;
      v_id := p_id;

      if not exists (select 1 from public.employees where role = 'admin' and active) then
        raise exception 'last_admin';
      end if;

      -- A changed PIN or a deactivation signs the employee out everywhere
      -- (except the admin's own current session).
      if v_pin is not null or coalesce(p_active, true) = false then
        delete from public.sessions
         where employee_id = p_id and token_hash <> public._hash_token(p_token);
      end if;

      perform public._audit(v_me, p_id, 'employee', 'update', null,
        jsonb_build_object('full_name', v_old.full_name, 'role', v_old.role, 'active', v_old.active),
        jsonb_build_object('full_name', v_name, 'role', p_role, 'active', coalesce(p_active, v_old.active),
                           'pin_changed', v_pin is not null));
    end if;
  exception when unique_violation then
    raise exception 'name_taken';
  end;

  return jsonb_build_object('ok', true, 'id', v_id);
end
$$;

-- Change history. Admin: everything (optionally for one employee).
-- Employee: only changes to their own days.
create or replace function public.get_audit_log(p_token text, p_employee_id uuid default null, p_limit integer default 50)
returns jsonb
language plpgsql stable security definer
set search_path = public, extensions, pg_temp
as $$
declare
  v_me     public.employees;
  v_filter uuid := p_employee_id;
begin
  v_me := public._current_employee(p_token);
  if v_me.role <> 'admin' then
    if v_filter is not null and v_filter <> v_me.id then
      raise exception 'forbidden';
    end if;
    v_filter := v_me.id;
  end if;

  return coalesce((select jsonb_agg(to_jsonb(x)) from (
    select al.id, al.changed_at, al.changed_by_name, al.employee_id, e.full_name as employee_name,
           al.entity, al.action, al.record_date, al.previous_value, al.new_value
      from public.audit_log al
      left join public.employees e on e.id = al.employee_id
     where (v_filter is null or al.employee_id = v_filter)
       and (v_me.role = 'admin' or al.entity = 'day')
     order by al.changed_at desc, al.id desc
     limit least(greatest(coalesce(p_limit, 50), 1), 500)) x), '[]'::jsonb);
end
$$;

-- Permissions: helpers are private, only the public API is callable.
revoke execute on function
  public._hash_token(text),
  public._current_employee(text),
  public._assert_admin(public.employees),
  public._assert_can_access(public.employees, uuid),
  public._day_state(uuid, date),
  public._audit(public.employees, uuid, text, text, date, jsonb, jsonb),
  public._parse_time(text)
from public, anon, authenticated;

revoke execute on function
  public.list_login_employees(),
  public.login(uuid, text),
  public.logout(text),
  public.get_app_data(text),
  public.save_day(text, uuid, date, text, text, text, text, text, text),
  public.reset_day(text, uuid, date),
  public.update_settings(text, text, text),
  public.save_employee(text, uuid, text, text, boolean, text),
  public.get_audit_log(text, uuid, integer)
from public;

grant execute on function
  public.list_login_employees(),
  public.login(uuid, text),
  public.logout(text),
  public.get_app_data(text),
  public.save_day(text, uuid, date, text, text, text, text, text, text),
  public.reset_day(text, uuid, date),
  public.update_settings(text, text, text),
  public.save_employee(text, uuid, text, text, boolean, text),
  public.get_audit_log(text, uuid, integer)
to anon, authenticated;
