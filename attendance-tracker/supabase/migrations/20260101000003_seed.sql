-- Initial data. PINs are hashed with bcrypt; the plain PIN is never stored.
-- Tomer's initial PIN is 0000 - change it in Settings after the first login.

insert into public.app_settings (id, year, default_start, default_end)
values (1, 2026, '08:00', '16:30')
on conflict (id) do nothing;

insert into public.employees (full_name, role, pin_hash)
select 'Tomer Cohen', 'admin', extensions.crypt('0000', extensions.gen_salt('bf', 10))
where not exists (select 1 from public.employees where lower(full_name) = 'tomer cohen');

insert into public.employees (full_name, role, pin_hash)
select 'Nir Ben Haim', 'employee', extensions.crypt('2222', extensions.gen_salt('bf', 10))
where not exists (select 1 from public.employees where lower(full_name) = 'nir ben haim');

insert into public.employees (full_name, role, pin_hash)
select 'Jonatan Ben Sudai', 'employee', extensions.crypt('1111', extensions.gen_salt('bf', 10))
where not exists (select 1 from public.employees where lower(full_name) = 'jonatan ben sudai');
