import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { BookOpen, Lock, Mail, UserPlus, ArrowRight, Loader2, Sparkles } from 'lucide-react'

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
    <div className="flex justify-center items-center min-h-screen p-4 sm:p-8 bg-slate-50 dark:bg-slate-950 relative overflow-hidden transition-colors">
      <div className="absolute top-[-10%] left-[-10%] w-96 h-96 bg-cyan-500/10 rounded-full blur-[100px] pointer-events-none"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-96 h-96 bg-indigo-500/10 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-[420px] bg-white dark:bg-slate-900 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-800 p-6 sm:p-8 relative z-10 transition-colors">
        
        <div className="flex justify-center mb-6">
          <div className="bg-gradient-to-tr from-cyan-500 to-indigo-600 p-4 rounded-2xl shadow-glow text-slate-950">
            <BookOpen className="w-8 h-8 text-slate-950" />
          </div>
        </div>
        
        <div className="text-center mb-8">
          <h2 className="text-2xl font-black text-slate-900 dark:text-white mb-1 tracking-tight">RTU DTR Portal</h2>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium flex items-center justify-center gap-1">
            {isLogin ? 'Welcome back, student!' : 'Create your OJT workspace'} <Sparkles className="w-3.5 h-3.5 text-cyan-500"/>
          </p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1.5 ml-1">Email Address</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Mail className="w-4 h-4"/></span>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-sm font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all min-w-0" 
                placeholder="student@rtu.edu.ph" />
            </div>
          </div>

          <div>
            <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1.5 ml-1">Password</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400"><Lock className="w-4 h-4"/></span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 outline-none text-sm font-bold text-slate-900 dark:text-white placeholder:text-slate-400 transition-all min-w-0" 
                placeholder="••••••••" />
            </div>
          </div>

          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-slate-900 dark:bg-cyan-500 hover:bg-cyan-600 dark:hover:bg-cyan-400 text-white dark:text-slate-950 font-bold py-4 px-4 rounded-xl shadow-lg transition-all active:scale-[0.98] disabled:opacity-50 mt-2">
            {loading ? <><Loader2 className="w-4 h-4 animate-spin"/> Processing...</> : isLogin ? <><ArrowRight className="w-4 h-4"/> Sign In to Workspace</> : <><UserPlus className="w-4 h-4"/> Create Account</>}
          </button>
        </form>
        
        <div className="mt-6 pt-5 border-t border-slate-100 dark:border-slate-800 text-center">
          <button onClick={() => setIsLogin(!isLogin)} className="text-xs font-bold text-slate-500 dark:text-slate-400 hover:text-cyan-600 dark:hover:text-cyan-400 transition-colors">
            {isLogin ? "Don't have an account? Sign up" : 'Already registered? Sign In'}
          </button>
        </div>
      </div>
    </div>
  )
}