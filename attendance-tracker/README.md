# Attendance & Vacation Tracker 2026

A simple, responsive attendance, vacation and sick-leave tracker for a small team, for calendar year 2026.
React + TypeScript + Vite, Supabase (Postgres) as the cloud database, installable as a PWA on iPhone, Android and desktop.

## How it works

- **Sunday–Thursday** are working days. Each one is automatically treated as **Present 08:00–16:30** (8.5 hours). Nobody needs to report a normal day.
- **Friday and Saturday** have no automatic hours. Employees can enter manual work on Friday. Only the administrator can enter Saturday work.
- **Holidays**: no automatic work, and vacation or sick leave cannot be recorded on them. They are not counted as working days.
- **Holiday eves**: automatic work until **14:00** (08:00–14:00 = 6 hours). The time can be changed in Settings.
- Only **exceptions** are stored in the database (vacation, sick leave, manual hours, notes). No rows are created for normal days.
- Partial leave is converted into a day fraction of the normal day: `hours / 8.5`. For example, 4.25 hours = 0.5 day.
- Statuses: Present, Vacation, Sick Leave, Partial Vacation, Partial Sick Leave, Manual Work, No Work.

### Users

| Employee | Role | Initial PIN |
|---|---|---|
| Tomer Cohen | Administrator | **0000** (change it in Settings after the first login) |
| Nir Ben Haim | Employee | 2222 |
| Jonatan Ben Sudai | Employee | 1111 |

The admin sees and edits everyone, reports, settings and the audit log. Employees see and edit only their own records.

## Run locally

Requires Node.js 18+.

```bash
cd attendance-tracker
npm install
npm run dev
```

Open http://localhost:5173.

**Without Supabase configured, the app runs in demo mode.** Data is kept in that browser's local storage only, and a yellow banner says so. This is useful for trying the app but is not shared between devices or users.

To use the real cloud database, configure Supabase (below), then:

```bash
cp .env.example .env      # fill in the two values
npm run dev
```

Other commands:

```bash
npm run build      # type-check and production build into dist/
npm run preview    # serve the production build locally
npm test           # unit tests: attendance calculations, reports, permissions (demo backend)
npm run test:db    # database tests against a real Postgres (see "Tests")
```

## Configure Supabase

1. Create a free project at https://supabase.com.
2. Apply the migrations in `supabase/migrations/` **in order**. Pick one option:
   - **SQL Editor**: open *SQL Editor* in the Supabase dashboard, paste each file and run it, in this order:
     1. `20260101000001_schema.sql` (tables)
     2. `20260101000002_functions.sql` (server-side API and permissions)
     3. `20260101000003_seed.sql` (settings and the three employees)
     4. `20260101000004_holidays.sql` (holidays, holiday eves and the 14:00 eve rule)
     5. `20260101000005_eve_half_vacation.sql` (vacation on a holiday eve = half a day)
   - **Supabase CLI**: `npx supabase link --project-ref <your-ref>` then `npx supabase db push`.
3. In *Project Settings → API*, copy the **Project URL** and the **anon / publishable key** into `.env`:
   ```
   VITE_SUPABASE_URL=https://xxxx.supabase.co
   VITE_SUPABASE_ANON_KEY=eyJ...
   ```
4. Restart `npm run dev`, then sign in as Tomer with PIN 0000 and change the PIN in **Settings**.

Supabase Auth is not used. Sign-in is by name + PIN, handled by the database functions.

## Deploy

`npm run build` produces a static site in `dist/`. Host it on any static host (Vercel, Netlify, Cloudflare Pages, GitHub Pages, Supabase Storage).
Set `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` as build environment variables. Routing uses `#/...` URLs, so no server rewrite rules are needed.
HTTPS is required for the PWA install prompt.

**Install on a phone**: open the site, then *Share → Add to Home Screen* on iPhone (Safari), or *Install app* on Android (Chrome). On desktop, use the install icon in the address bar.

## Security

- PINs are stored as **bcrypt hashes** (`pgcrypto`), never as plain text. Session tokens are stored as SHA-256 hashes.
- **All tables have Row Level Security enabled with no policies, and all table privileges are revoked** from the public `anon` / `authenticated` roles. The browser cannot read or write any table directly.
- The browser can only call a fixed set of `SECURITY DEFINER` functions. Each one resolves the session token to an employee and checks the role **on the server**:
  - employees can read and change only their own records, so changing a URL or editing a request returns `forbidden`
  - settings, employee management and the full audit log are admin only
  - Saturday work is admin only
- 5 wrong PINs lock the employee for 5 minutes. Changing a PIN or deactivating an employee ends their sessions.
- Every change writes an **audit log** row: who made it, when, previous value and new value.
- The login screen lists active employee names, which is public by design for a small team.

## Database

| Table | Purpose |
|---|---|
| `employees` | name, role (`admin`/`employee`), bcrypt PIN hash, active flag, lockout counters |
| `attendance_records` | one row per employee-day **exception**: status, work start/end, notes |
| `leave_records` | vacation/sick for that day: full day or start/end, hours |
| `app_settings` | single row: year, default start and end time, holiday eve end time |
| `holidays` | date, name, type (`holiday` or `eve`), editable by the admin |
| `audit_log` | who/when/previous value/new value for every change |
| `sessions` | hashed login tokens (30 days) |

Server API (RPC functions): `list_login_employees`, `login`, `logout`, `get_app_data`, `save_day`, `reset_day`, `update_settings`, `save_employee`, `get_audit_log`, `save_holiday`, `delete_holiday`.

## Tests

- `npm test` runs 51 unit tests:
  - date rules for all of 2026
  - automatic days
  - full and partial vacation
  - full and partial sick leave
  - Friday manual work
  - holidays and holiday eves
  - monthly and annual totals
  - CSV export
  - permissions in the demo backend
- `npm run test:db` runs 29 tests against **real Postgres** with the actual migrations. It covers:
  - PIN hashing and lockout
  - no direct table access for the public role
  - employee isolation
  - admin-only functions
  - Saturday rule
  - leave records and audit log
  - holidays, holiday eves and admin-only holiday management
  - adding employees, PIN change, deactivation
  - last-admin protection

  It creates and drops a temporary database, so point it at any Postgres where you may create databases:
  ```bash
  TEST_DATABASE_URL=postgres://user:pass@localhost:5432/postgres npm run test:db
  ```

## Project structure

```
attendance-tracker/
├── index.html
├── package.json
├── vite.config.ts            # Vite + PWA (manifest, service worker)
├── .env.example
├── public/                   # app icons
├── supabase/migrations/      # schema, server functions, seed data
├── src/
│   ├── main.tsx / App.tsx    # entry and routes
│   ├── styles.css            # responsive styles
│   ├── api/                  # backend: Supabase RPC client + in-browser demo backend
│   ├── context/              # session and data state
│   ├── lib/
│   │   ├── dates.ts          # 2026 calendar and day-of-week rules
│   │   ├── attendance.ts     # the attendance rules and calculations
│   │   ├── reports.ts        # monthly and annual report tables
│   │   └── export.ts         # PDF and CSV export
│   ├── components/           # calendar, daily record dialog, layout, selectors
│   └── pages/                # Dashboard, Calendar, Report Leave, Reports, Settings, Login
├── tests/                    # unit tests
└── db-tests/                 # Postgres tests for the SQL functions
```

## Known limitations

- Holidays seeded for 2026:

  | Date | Holiday | Type |
  |---|---|---|
  | 01/04 Wed | Erev Pesach | eve |
  | 02/04 Thu | Pesach | holiday |
  | 07/04 Tue | Erev Seventh Day of Pesach | eve |
  | 08/04 Wed | Seventh Day of Pesach | holiday |
  | 22/04 Wed | Independence Day | holiday |
  | 21/05 Thu | Erev Shavuot | eve |
  | 22/05 Fri | Shavuot | holiday |
  | 11/09 Fri | Erev Rosh Hashana | eve |
  | 12/09 Sat | Rosh Hashana | holiday |
  | 13/09 Sun | Rosh Hashana (second day) | holiday |
  | 20/09 Sun | Erev Yom Kippur | eve |
  | 21/09 Mon | Yom Kippur | holiday |
  | 25/09 Fri | Erev Sukkot | eve |
  | 26/09 Sat | Sukkot | holiday |
  | 02/10 Fri | Erev Simchat Torah | eve |
  | 03/10 Sat | Simchat Torah | holiday |

  Chol HaMoed, Memorial Day eve and election day are not included. The admin can add them in **Settings → Holidays**.
- A full vacation day on a holiday eve counts as **half a vacation day** (4.25 hours). A full sick day on an eve counts as 1 sick day (6 hours). Partial leave is always a fraction of the normal 8.5-hour day.
- Changing the default working hours in Settings also changes how past automatic days are counted, because normal days are not stored.
- Demo mode is per browser and only for trying the app.
