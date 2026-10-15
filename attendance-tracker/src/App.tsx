import { HashRouter, Navigate, Route, Routes } from 'react-router-dom';
import { Layout } from './components/Layout';
import { AppProvider } from './context/AppContext';
import { CalendarPage } from './pages/CalendarPage';
import { DashboardPage } from './pages/DashboardPage';
import { EmployeePage } from './pages/EmployeePage';
import { LoginPage } from './pages/LoginPage';
import { ReportLeavePage } from './pages/ReportLeavePage';
import { ReportsPage } from './pages/ReportsPage';
import { SettingsPage } from './pages/SettingsPage';

export default function App() {
  return (
    <AppProvider
      signedIn={
        <HashRouter>
          <Routes>
            <Route element={<Layout />}>
              <Route index element={<DashboardPage />} />
              <Route path="calendar" element={<CalendarPage />} />
              <Route path="report" element={<ReportLeavePage />} />
              <Route path="reports" element={<ReportsPage />} />
              <Route path="settings" element={<SettingsPage />} />
              <Route path="employee/:id" element={<EmployeePage />} />
              <Route path="*" element={<Navigate to="/" replace />} />
            </Route>
          </Routes>
        </HashRouter>
      }
    >
      {({ status, login }) => <LoginPage loading={status === 'loading'} onLogin={login} />}
    </AppProvider>
  );
}
