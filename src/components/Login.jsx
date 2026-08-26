import React, { useState } from 'react';
import { Lock } from 'lucide-react';
import { signInWithEmailAndPassword } from "firebase/auth";
import { auth } from '../firebase'; 

export default function Login() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError("");
    try { 
      await signInWithEmailAndPassword(auth, email, password); 
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
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" />
          </div>
          <div>
            <label className="block text-xs font-bold text-gray-500 uppercase tracking-wider mb-2">Password</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-4 py-3.5 bg-gray-50 border border-gray-200 rounded-xl text-sm font-medium focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all" />
          </div>
          {loginError && <p className="text-red-500 text-xs font-bold text-center bg-red-50 py-2 rounded-lg">{loginError}</p>}
          <button type="submit" className="w-full py-4 bg-gray-900 hover:bg-blue-600 text-white font-bold rounded-xl shadow-md transition-colors duration-300">Access Dashboard</button>
        </form>
      </div>
    </div>
  );
}