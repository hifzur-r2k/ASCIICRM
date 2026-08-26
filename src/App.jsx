import React, { useState, useEffect } from 'react';
import { onAuthStateChanged, signOut } from "firebase/auth";
import { auth } from './firebase';

// Import our new modules
import Login from './components/Login';
import AdminDashboard from './components/AdminDashboard';
import SupervisorPortal from './components/SupervisorPortal';

export default function App() {
  const [user, setUser] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);

  // Listen for Firebase login state
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setAuthLoading(false);
    });
    return () => unsubscribe();
  }, []);

  const handleLogout = async () => {
    await signOut(auth);
  };

  if (authLoading) {
    return <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] text-gray-500 font-bold">Initializing Secure Environment...</div>;
  }

  // If no one is logged in, show the Login screen
  if (!user) {
    return <Login />;
  }

  // --- THE TRAFFIC COP LOGIC ---
  // Change this to your ACTUAL admin email address!
  const MASTER_ADMIN_EMAIL = "hifzur.r2k@gmail.com";

  // Security Wall: Only the exact Master Admin email gets the Dashboard code
  if (user.email === MASTER_ADMIN_EMAIL) {
    return <AdminDashboard currentUser={user} onLogout={handleLogout} />;
  } 
  
  // Everyone else (Supervisors) gets locked into the Field Portal
  else {
    return <SupervisorPortal currentUser={user} onLogout={handleLogout} />;
  }
}