-- A full vacation day on a holiday eve counts as half a vacation day
-- (half of the normal 8.5-hour day = 4.25 hours). Sick leave on an eve is
-- unchanged: a full sick day of the short eve length.

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
  v_normal   numeric;
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
  v_normal := extract(epoch from (v_settings.default_end - v_settings.default_start)) / 3600;
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
        case when p_status = 'vacation' and v_kind = 'eve' then v_normal / 2
             when p_status in ('vacation', 'sick') then v_full
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

grant execute on function public.save_day(text, uuid, date, text, text, text, text, text, text) to anon, authenticated;
