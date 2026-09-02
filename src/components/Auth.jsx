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
    <div className="flex justify-center items-center min-h-screen p-4 sm:p-8 bg-dark-950 relative overflow-hidden">
      {/* Modern Gradient Orbs */}
      <div className="absolute top-[-10%] left-[-10%] w-96 h-96 bg-blue-600/15 rounded-full blur-[100px] pointer-events-none"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-96 h-96 bg-indigo-600/15 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-[420px] bg-dark-900/90 backdrop-blur-2xl rounded-3xl shadow-2xl border border-slate-800 p-8 relative z-10">
        
        <div className="flex items-center justify-between mb-8 pb-5 border-b border-slate-800">
           <div className="flex items-center gap-3">
              <div className="bg-blue-500/10 p-3 rounded-2xl border border-blue-500/20 text-blue-400">
                <BookOpen className="w-6 h-6" />
              </div>
              <div>
                <span className="text-[10px] font-bold text-blue-400 uppercase tracking-widest block">Rizal Technological University</span>
                <h1 className="text-sm font-black text-white tracking-wide">OJT DTR PORTAL</h1>
              </div>
           </div>
        </div>

        <div className="mb-6">
          <h2 className="text-xl font-black text-white mb-1">{isLogin ? 'Welcome Back' : 'Create Student Account'}</h2>
          <p className="text-xs text-slate-400">Sign in to manage and export your daily time records.</p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 ml-1">Email Address</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-500"><Mail className="w-4 h-4"/></span>
              <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 outline-none text-sm font-bold text-white placeholder:text-slate-600 transition-all" 
                placeholder="student@rtu.edu.ph" />
            </div>
          </div>

          <div>
            <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1.5 ml-1">Password</label>
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-500"><Lock className="w-4 h-4"/></span>
              <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
                className="w-full pl-10 pr-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 outline-none text-sm font-bold text-white placeholder:text-slate-600 transition-all" 
                placeholder="••••••••" />
            </div>
          </div>

          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-500 text-white font-bold py-4 px-4 rounded-xl shadow-lg shadow-blue-600/25 transition-all active:scale-[0.98] disabled:opacity-50 mt-2">
            {loading ? <><Loader2 className="w-4 h-4 animate-spin"/> Processing...</> : isLogin ? <><ArrowRight className="w-4 h-4"/> Sign In</> : <><UserPlus className="w-4 h-4"/> Register Account</>}
          </button>
        </form>
        
        <div className="mt-6 pt-5 border-t border-slate-800 text-center">
          <button onClick={() => setIsLogin(!isLogin)} className="text-xs font-bold text-slate-400 hover:text-blue-400 transition-colors">
            {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Sign in'}
          </button>
        </div>
      </div>
    </div>
  )
}