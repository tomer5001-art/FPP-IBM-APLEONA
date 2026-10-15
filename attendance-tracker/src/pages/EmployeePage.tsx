import { Link, Navigate, useParams } from 'react-router-dom';
import { ChangeList } from '../components/ChangeList';
import { useApp } from '../context/AppContext';
import { EmployeeOverview } from './EmployeeOverview';

/** Admin view of one employee's full record. Non-admins only ever see themselves. */
export function EmployeePage() {
  const { id = '' } = useParams();
  const { data, isAdmin, employeeById } = useApp();
  if (!isAdmin || !employeeById(id)) return <Navigate to="/" replace />;
  return (
    <div className="stack">
      <Link to="/" className="link">
        ← Back to dashboard
      </Link>
      <EmployeeOverview employeeId={id} />
      <section className="card">
        <div className="card-head">
          <h2>Change history</h2>
          <Link className="btn btn-small" to={`/reports?employee=${id}`}>
            Monthly report
          </Link>
        </div>
        <ChangeList employeeId={id} key={data.records.length} />
      </section>
    </div>
  );
}
