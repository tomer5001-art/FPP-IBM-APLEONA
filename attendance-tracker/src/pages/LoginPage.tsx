import { useEffect, useState } from 'react';
import { api } from '../api';
import { describeError } from '../context/AppContext';

export function LoginPage({ onLogin, loading }: { onLogin(token: string): void; loading: boolean }) {
  const [people, setPeople] = useState<{ id: string; fullName: string }[]>([]);
  const [employeeId, setEmployeeId] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api
      .listLoginEmployees()
      .then((list) => {
        setPeople(list);
        if (list.length) setEmployeeId(list[0].id);
      })
      .catch((e) => setError(describeError(e)));
  }, []);

  async function submit(e: React.FormEvent) {
    e.preventDefault();
    setError('');
    setBusy(true);
    try {
      const res = await api.login(employeeId, pin);
      onLogin(res.token);
    } catch (err) {
      setError(describeError(err));
      setPin('');
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="login">
      <form className="card login-card" onSubmit={submit}>
        <img src="icon.svg" alt="" width={56} height={56} />
        <h1>Attendance 2026</h1>
        <p className="muted">Sign in with your name and PIN</p>
        {loading ? (
          <p className="muted">Loading…</p>
        ) : (
          <>
            <label className="field">
              <span>Name</span>
              <select value={employeeId} onChange={(e) => setEmployeeId(e.target.value)} required>
                {people.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.fullName}
                  </option>
                ))}
              </select>
            </label>
            <label className="field">
              <span>PIN</span>
              <input
                type="password"
                inputMode="numeric"
                autoComplete="current-password"
                pattern="[0-9]*"
                maxLength={8}
                value={pin}
                onChange={(e) => setPin(e.target.value.replace(/\D/g, ''))}
                required
                autoFocus
              />
            </label>
            {error && <p className="alert">{error}</p>}
            <button type="submit" className="btn btn-primary btn-block" disabled={busy || !employeeId || pin.length < 4}>
              {busy ? 'Signing in…' : 'Sign in'}
            </button>
            {api.mode === 'demo' && <p className="muted small">Demo mode · Tomer 0000 · Nir 2222 · Jonatan 1111</p>}
          </>
        )}
      </form>
    </div>
  );
}
