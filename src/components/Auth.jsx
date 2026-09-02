import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { ShieldCheck, Terminal, Lock, UserPlus, ArrowRight, Loader2, Cpu, ShieldAlert } from 'lucide-react'

export default function Auth() {
  const [loading, setLoading] = useState(false)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [isLogin, setIsLogin] = useState(true)

  const handleAuth = async (e) => {
    e.preventDefault()
    setLoading(true)
    try {
      if (isLogin) {
        const { error } = await supabase.auth.signInWithPassword({ email, password })
        if (error) throw error
      } else {
        const { error } = await supabase.auth.signUp({ email, password })
        if (error) throw error
        alert('Clearance granted! You can now authenticate.')
      }
    } catch (error) {
      alert(error.error_description || error.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="flex justify-center items-center min-h-screen p-4 sm:p-6 bg-slate-50 dark:bg-slate-950 relative overflow-hidden font-mono transition-colors duration-300">
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#33415510_1px,transparent_1px),linear-gradient(to_bottom,#33415510_1px,transparent_1px)] bg-[size:3rem_3rem] pointer-events-none"></div>
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 -translate-y-1/2 w-80 h-80 bg-cyan-500/10 dark:bg-cyan-500/10 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-[400px] bg-white dark:bg-slate-900 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-800 p-6 sm:p-8 relative z-10 transition-colors">
        
        <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100 dark:border-slate-800">
           <div className="flex items-center gap-2.5">
              <div className="bg-cyan-500/10 p-2 rounded-xl border border-cyan-500/30 text-cyan-500 dark:text-cyan-400">
                <ShieldCheck className="w-5 h-5" />
              </div>
              <div>
                <span className="text-[9px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest block">RTU Cyber Defense</span>
                <h1 className="text-xs font-black text-slate-900 dark:text-white tracking-wider">SECURE DTR PORTAL</h1>
              </div>
           </div>
           <Cpu className="w-4 h-4 text-slate-400 dark:text-slate-600 animate-pulse" />
        </div>

        <div className="mb-5">
          <h2 className="text-lg font-bold text-slate-900 dark:text-white mb-1">{isLogin ? 'Operative Login' : 'Security Clearance'}</h2>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-sans">Authenticate to access encrypted timesheet network.</p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1">Secure Email</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400"><Terminal className="w-3.5 h-3.5"/></span>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
                className="w-full pl-9 pr-3 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-xs font-sans font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all" 
                placeholder="operative@rtu.edu.ph" />
            </div>
          </div>

          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1">Encryption Key</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400"><Lock className="w-3.5 h-3.5"/></span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
                className="w-full pl-9 pr-3 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-xs font-sans font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all" 
                placeholder="••••••••" />
            </div>
          </div>

          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950 font-black py-3.5 px-4 rounded-xl shadow-md shadow-cyan-500/20 transition-all active:scale-[0.98] disabled:opacity-50 mt-2 font-sans tracking-wide text-xs">
            {loading ? <><Loader2 className="w-4 h-4 animate-spin"/> Authenticating...</> : isLogin ? <><ArrowRight className="w-4 h-4"/> Initialize Session</> : <><UserPlus className="w-4 h-4"/> Request Clearance</>}
          </button>
        </form>
        
        <div className="mt-5 pt-4 border-t border-slate-100 dark:border-slate-800 text-center font-sans">
          <button onClick={() => setIsLogin(!isLogin)} className="text-xs font-bold text-slate-500 dark:text-slate-400 hover:text-cyan-600 dark:hover:text-cyan-400 transition-colors">
            {isLogin ? "Need clearance? Register account" : 'Already cleared? Authenticate'}
          </button>
        </div>
      </div>
    </div>
  )
}