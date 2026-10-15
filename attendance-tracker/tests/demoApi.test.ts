import { beforeEach, describe, expect, it } from 'vitest';
import type { Api } from '../src/api/api';
import { createDemoApi, type KeyValueStore } from '../src/api/demoApi';

function memoryStore(): KeyValueStore {
  const m = new Map<string, string>();
  return { getItem: (k) => m.get(k) ?? null, setItem: (k, v) => void m.set(k, v) };
}

let api: Api;
let ids: Record<string, string>;

async function signIn(name: string, pin: string) {
  return (await api.login(ids[name], pin)).token;
}

beforeEach(async () => {
  api = createDemoApi(memoryStore());
  const people = await api.listLoginEmployees();
  ids = Object.fromEntries(people.map((p) => [p.fullName, p.id]));
});

describe('demo backend permissions', () => {
  it('seeds the three employees', () => {
    expect(Object.keys(ids)).toEqual(['Tomer Cohen', 'Nir Ben Haim', 'Jonatan Ben Sudai']);
  });

  it('rejects a wrong PIN', async () => {
    await expect(api.login(ids['Nir Ben Haim'], '1111')).rejects.toMatchObject({ code: 'invalid_credentials' });
  });

  it('employee sees only their own data', async () => {
    const nir = await signIn('Nir Ben Haim', '2222');
    const d = await api.getAppData(nir);
    expect(d.employees.map((e) => e.fullName)).toEqual(['Nir Ben Haim']);
  });

  it('employee cannot edit another employee', async () => {
    const nir = await signIn('Nir Ben Haim', '2222');
    await expect(
      api.saveDay(nir, { employeeId: ids['Jonatan Ben Sudai'], date: '2026-01-05', status: 'vacation' }),
    ).rejects.toMatchObject({ code: 'forbidden' });
  });

  it('employee cannot use admin functions', async () => {
    const jon = await signIn('Jonatan Ben Sudai', '1111');
    await expect(api.updateSettings(jon, '07:00', '15:00')).rejects.toMatchObject({ code: 'forbidden' });
    await expect(api.saveEmployee(jon, { id: null, fullName: 'X', role: 'admin', active: true, pin: '1234' })).rejects.toMatchObject({ code: 'forbidden' });
  });

  it('admin can edit everyone and the audit log records it', async () => {
    const tomer = await signIn('Tomer Cohen', '0000');
    await api.saveDay(tomer, { employeeId: ids['Jonatan Ben Sudai'], date: '2026-01-03', status: 'manual', startTime: '09:00', endTime: '13:00' });
    const d = await api.getAppData(tomer);
    expect(d.employees).toHaveLength(3);
    expect(d.records).toHaveLength(1);
    const log = await api.getAuditLog(tomer, null);
    expect(log[0]).toMatchObject({ changedByName: 'Tomer Cohen', action: 'update', recordDate: '2026-01-03', previousValue: null });
  });

  it('saving the default day stores nothing', async () => {
    const nir = await signIn('Nir Ben Haim', '2222');
    await api.saveDay(nir, { employeeId: ids['Nir Ben Haim'], date: '2026-01-05', status: 'present', startTime: '08:00', endTime: '16:30' });
    expect((await api.getAppData(nir)).records).toHaveLength(0);
  });
});
