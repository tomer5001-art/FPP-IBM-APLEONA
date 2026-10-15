import { NavLink, Outlet } from 'react-router-dom';
import { api } from '../api';
import { useApp } from '../context/AppContext';

const ICONS: Record<string, string> = {
  Dashboard: '▦',
  Calendar: '▤',
  'Report Leave': '✎',
  Reports: '▥',
  Settings: '⚙',
};

export function Layout() {
  const { data, isAdmin, logout } = useApp();
  const links = [
    { to: '/', label: 'Dashboard' },
    { to: '/calendar', label: 'Calendar' },
    { to: '/report', label: 'Report Leave' },
    { to: '/reports', label: 'Reports' },
    ...(isAdmin ? [{ to: '/settings', label: 'Settings' }] : []),
  ];
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <img src="icon.svg" alt="" width={28} height={28} />
          <span>Attendance 2026</span>
        </div>
        <nav className="nav-desktop">
          {links.map((l) => (
            <NavLink key={l.to} to={l.to} end={l.to === '/'}>
              {l.label}
            </NavLink>
          ))}
        </nav>
        <div className="user">
          <span className="user-name">
            {data.me.fullName}
            {isAdmin && <span className="tag">Admin</span>}
          </span>
          <button type="button" className="btn btn-small" onClick={logout}>
            Sign out
          </button>
        </div>
      </header>
      {api.mode === 'demo' && (
        <div className="demo-banner">Demo mode: data is stored only in this browser. Configure Supabase to share data across devices.</div>
      )}
      <main className="content">
        <Outlet />
      </main>
      <nav className="nav-mobile">
        {links.map((l) => (
          <NavLink key={l.to} to={l.to} end={l.to === '/'}>
            <span className="nav-icon" aria-hidden>
              {ICONS[l.label]}
            </span>
            <span>{l.label}</span>
          </NavLink>
        ))}
      </nav>
    </div>
  );
}
