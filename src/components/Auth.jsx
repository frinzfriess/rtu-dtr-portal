import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { ShieldCheck, Lock, Mail, UserPlus, ArrowRight, Loader2, Cpu } from 'lucide-react'

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
        alert('Registration successful! You can now log in.')
      }
    } catch (error) {
      alert(error.error_description || error.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="flex justify-center items-center min-h-screen p-4 sm:p-8 bg-slate-50 dark:bg-slate-950 relative overflow-hidden transition-colors font-mono">
      <div className="absolute inset-0 bg-[linear-gradient(to_right,#33415510_1px,transparent_1px),linear-gradient(to_bottom,#33415510_1px,transparent_1px)] bg-[size:3rem_3rem] pointer-events-none"></div>
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 -translate-y-1/2 w-96 h-96 bg-cyan-500/10 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-[420px] bg-white dark:bg-slate-900 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-800 p-8 relative z-10 transition-colors">
        <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100 dark:border-slate-800">
           <div className="flex items-center gap-2.5">
              <div className="bg-cyan-500/10 p-2.5 rounded-xl border border-cyan-500/30 text-cyan-600 dark:text-cyan-400">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <span className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest block">RTU OJT Portal</span>
                <h1 className="text-sm font-black text-slate-900 dark:text-white tracking-wider">SECURE DTR LOGIN</h1>
              </div>
           </div>
           <Cpu className="w-5 h-5 text-slate-400 dark:text-slate-600 animate-pulse" />
        </div>

        <div className="mb-6">
          <h2 className="text-xl font-bold text-slate-900 dark:text-white mb-1">{isLogin ? 'Student Authentication' : 'Create Account'}</h2>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-sans">Enter your credentials to access your DTR records.</p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 ml-1">Email Address</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Mail className="w-4 h-4"/></span>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-sm font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all min-w-0" 
                placeholder="student@rtu.edu.ph" />
            </div>
          </div>

          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 ml-1">Password</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Lock className="w-4 h-4"/></span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-sm font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all min-w-0" 
                placeholder="••••••••" />
            </div>
          </div>

          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950 font-black py-4 px-4 rounded-xl shadow-lg shadow-cyan-500/20 transition-all active:scale-[0.98] disabled:opacity-50 mt-2 font-sans tracking-wide text-xs">
            {loading ? <><Loader2 className="w-4 h-4 animate-spin"/> Authenticating...</> : isLogin ? <><ArrowRight className="w-4 h-4"/> Sign In</> : <><UserPlus className="w-4 h-4"/> Register</>}
          </button>
        </form>
        
        <div className="mt-6 pt-5 border-t border-slate-100 dark:border-slate-800 text-center font-sans">
          <button onClick={() => setIsLogin(!isLogin)} className="text-xs font-bold text-slate-500 dark:text-slate-400 hover:text-cyan-600 dark:hover:text-cyan-400 transition-colors">
            {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Sign in'}
          </button>
        </div>
      </div>
    </div>
  )
}