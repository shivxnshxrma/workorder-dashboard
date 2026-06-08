'use client';

import React, { useState, useEffect } from 'react';
import { Mail, Lock, LogIn, AlertCircle, Database } from 'lucide-react';
import DashboardClient from '@/components/DashboardClient';
import './login.css';
import './dashboard.css';

export default function Home() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [userEmail, setUserEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  // Check login status on load
  useEffect(() => {
    const checkSession = () => {
      const cookies = document.cookie.split(';');
      const sessionCookie = cookies.find((c) => c.trim().startsWith('dashboard_session='));
      const savedEmail = localStorage.getItem('soteria_user_email');
      
      if (sessionCookie && savedEmail) {
        setIsAuthenticated(true);
        setUserEmail(savedEmail);
      } else {
        setIsAuthenticated(false);
      }
    };
    checkSession();
  }, []);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!email || !password) {
      setError('Please fill in all fields.');
      return;
    }

    setLoading(true);
    setError('');

    try {
      const response = await fetch('/api/auth', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email, password }),
      });

      const data = await response.json();

      if (response.ok && data.success) {
        localStorage.setItem('soteria_user_email', data.user.email);
        setUserEmail(data.user.email);
        setIsAuthenticated(true);
      } else {
        setError(data.error || 'Invalid credentials. Please try again.');
      }
    } catch (err) {
      setError('A connection error occurred. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = () => {
    // Delete session cookie
    document.cookie = 'dashboard_session=; path=/; expires=Thu, 01 Jan 1970 00:00:01 GMT;';
    localStorage.removeItem('soteria_user_email');
    setIsAuthenticated(false);
    setEmail('');
    setPassword('');
  };

  // Prevent flash before checking authentication
  if (isAuthenticated === null) {
    return (
      <div style={{ display: 'flex', minHeight: '100vh', alignItems: 'center', justifyContent: 'center', backgroundColor: '#0a0c10', color: '#f8fafc' }}>
        <div className="loading-spinner"></div>
      </div>
    );
  }

  if (isAuthenticated) {
    return <DashboardClient userEmail={userEmail} onLogout={handleLogout} />;
  }

  return (
    <div className="login-container animated-fade-in">
      <div className="login-glass-card glow-effect">
        <div className="login-header">
          <div className="login-logo">
            <Database size={24} />
          </div>
          <h2 className="login-title">Soteria Control Center</h2>
          <p className="login-subtitle">Sign in to manage bulk work orders</p>
        </div>

        <form className="login-form" onSubmit={handleLogin}>
          <div className="input-group">
            <label className="input-label" htmlFor="email-input">EMAIL ADDRESS</label>
            <div className="input-field-wrapper">
              <Mail className="input-icon" size={16} />
              <input
                id="email-input"
                type="email"
                className="input-field transition-all"
                placeholder="admin@soteria.in"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                disabled={loading}
              />
            </div>
          </div>

          <div className="input-group">
            <label className="input-label" htmlFor="password-input">PASSWORD</label>
            <div className="input-field-wrapper">
              <Lock className="input-icon" size={16} />
              <input
                id="password-input"
                type="password"
                className="input-field transition-all"
                placeholder="••••••••••••"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                disabled={loading}
              />
            </div>
          </div>

          {error && (
            <div className="auth-error">
              <AlertCircle size={16} style={{ flexShrink: 0 }} />
              <span>{error}</span>
            </div>
          )}

          <button className="login-btn transition-all" type="submit" disabled={loading}>
            {loading ? (
              <span className="loading-spinner"></span>
            ) : (
              <>
                <LogIn size={18} />
                <span>Access Dashboard</span>
              </>
            )}
          </button>
        </form>
      </div>
    </div>
  );
}
