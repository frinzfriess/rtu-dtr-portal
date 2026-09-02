import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { ShieldCheck, Lock, Mail, UserPlus, ArrowRight, Loader2 } from 'lucide-react'

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
    <div className="flex justify-center items-center min-h-screen p-4 bg-slate-100 dark:bg-slate-950 relative overflow-hidden font-mono">
      <div className="w-full max-w-[420px] bg-white dark:bg-slate-900 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-800 p-6 sm:p-8 relative z-10">
        <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100 dark:border-slate-800">
           <div className="flex items-center gap-2.5">
              <div className="bg-cyan-500/10 p-2.5 rounded-xl border border-cyan-500/30 text-cyan-600 dark:text-cyan-400">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <span className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest block">RTU OJT Portal</span>
                <h1 className="text-sm font-black text-slate-900 dark:text-white tracking-wider">SECURE LOGIN</h1>
              </div>
           </div>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 ml-1">Email</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Mail className="w-4 h-4"/></span>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none text-sm font-bold text-slate-900 dark:text-white" 
                placeholder="student@rtu.edu.ph" />
            </div>
          </div>

          <div>
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 ml-1">Password</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Lock className="w-4 h-4"/></span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none text-sm font-bold text-slate-900 dark:text-white" 
                placeholder="••••••••" />
            </div>
          </div>

          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black py-3.5 px-4 rounded-xl shadow-lg transition-all text-xs">
            {loading ? <Loader2 className="w-4 h-4 animate-spin"/> : <ArrowRight className="w-4 h-4"/>} {isLogin ? 'Sign In' : 'Register'}
          </button>
        </form>
        
        <div className="mt-6 pt-5 border-t border-slate-100 dark:border-slate-800 text-center font-sans">
          <button onClick={() => setIsLogin(!isLogin)} className="text-xs font-bold text-slate-500 dark:text-slate-400 hover:text-cyan-500">
            {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Sign in'}
          </button>
        </div>
      </div>
    </div>
  )
}