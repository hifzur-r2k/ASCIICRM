import React, { useState } from 'react';
import { Lock, Eye, EyeOff } from 'lucide-react';
import { signInWithEmailAndPassword } from "firebase/auth";
import { collection, addDoc } from "firebase/firestore";
import { auth, db } from '../firebase'; 

// Helper function to capture session metadata
const recordAdminSession = async (userEmail) => {
  const agent = navigator.userAgent;
  const isMobile = /Mobile|Android|iP(ad|hone)/i.test(agent);
  const deviceType = isMobile ? 'Mobile' : 'Desktop PC';

  let ipData = { ip: "Unavailable", city: "Unknown", region: "Unknown", country_name: "Unknown" };

  try {
    const res = await fetch('https://ipapi.co/json/');
    if (res.ok) {
      ipData = await res.json();
    }
  } catch (netErr) {
    // Continues cleanly if an ad-blocker or browser shield stops the IP service
  }

  try {
    await addDoc(collection(db, "admin_sessions"), {
      email: userEmail,
      ip: ipData.ip || "Unavailable",
      city: ipData.city || "Unknown",
      region: ipData.region || "Unknown",
      country: ipData.country_name || "Unknown",
      device: deviceType,
      browserAgent: agent,
      timestamp: Date.now(),
      loginTime: new Date().toLocaleString('en-IN')
    });
  } catch (dbErr) {
    console.error("Firestore session record failed:", dbErr);
  }
};

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loginError, setLoginError] = useState("");

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError("");
    try { 
      const cred = await signInWithEmailAndPassword(auth, email, password);
      // Wait for session record to save before the dashboard view loads
      await recordAdminSession(cred.user?.email || email);
    } catch (err) { 
      setLoginError("Invalid email or password. Please try again."); 
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-[#f8fafc] px-4 font-sans selection:bg-blue-100 selection:text-blue-900">
      <div className="w-full max-w-md bg-white p-8 md:p-10 rounded-[2rem] shadow-xl border border-gray-100">
        <div className="flex flex-col items-center mb-8 text-center">
          <div className="bg-blue-600 p-4 rounded-2xl shadow-lg shadow-blue-600/20 mb-5">
            <Lock className="w-8 h-8 text-white" />
          </div>
          <h1 className="text-3xl font-black text-gray-900 tracking-tight">CRM_FIX</h1>
          <p className="text-sm font-bold text-gray-400 uppercase tracking-widest mt-1">Authorized Access Only</p>
        </div>
        <form onSubmit={handleLogin} className="space-y-5">
          <div>
            <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Email</label>
            <input 
              type="email" 
              value={email} 
              onChange={(e) => setEmail(e.target.value)} 
              required 
              className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" 
            />
          </div>
          
          <div>
            <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Password</label>
            <div className="relative w-full">
              <input 
                type={showPassword ? "text" : "password"} 
                value={password} 
                onChange={(e) => setPassword(e.target.value)} 
                required 
                className="w-full pl-4 pr-12 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" 
              />
              <button 
                type="button" 
                onClick={() => setShowPassword(!showPassword)} 
                className="absolute right-4 top-1/2 -translate-y-1/2 text-gray-400 hover:text-blue-600 outline-none transition-colors"
              >
                {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
              </button>
            </div>
          </div>

          {loginError && <p className="text-red-500 text-xs font-bold text-center bg-red-50 py-2 rounded-lg">{loginError}</p>}
          <button type="submit" className="w-full py-4 bg-gray-900 hover:bg-blue-600 text-white font-bold rounded-xl shadow-md transition-colors duration-300">Access Dashboard</button>
        </form>
      </div>
    </div>
  );
}