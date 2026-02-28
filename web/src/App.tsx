import { useEffect } from 'react';
import { Route, Switch, useLocation } from 'wouter';
import Login from './pages/Login';
import Dashboard from './pages/Dashboard';
import OAuthCallback from './pages/OAuthCallback';

// Protect routes that require authentication
const ProtectedRoute = ({ component: Component }: { component: any }) => {
  const [, setLocation] = useLocation();
  const token = localStorage.getItem('auth_token');

  useEffect(() => {
    if (!token) {
      setLocation('/login');
    }
  }, [token, setLocation]);

  if (!token) {
    return null; // Will redirect in useEffect
  }

  return <Component />;
};

function App() {
  return (
    <Switch>
      <Route path="/login" component={Login} />
      <Route path="/login/success" component={OAuthCallback} />
      <Route path="/">
        <ProtectedRoute component={Dashboard} />
      </Route>
      {/* Fallback route */}
      <Route>
        <ProtectedRoute component={Dashboard} />
      </Route>
    </Switch>
  );
}

export default App;
