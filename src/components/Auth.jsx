import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { BookOpen, ArrowRight, UserPlus, Sparkles, Loader2 } from 'lucide-react'

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
    <div className="flex justify-center items-center min-h-[100dvh] p-4 sm:p-8 bg-[url('https://www.transparenttextures.com/patterns/cubes.png')] bg-slate-50 relative overflow-hidden">
      {/* Decorative Orbs */}
      <div className="absolute top-[-15%] left-[-10%] w-96 h-96 bg-indigo-400 rounded-full mix-blend-multiply filter blur-[100px] opacity-40 animate-pulse"></div>
      <div className="absolute bottom-[-15%] right-[-10%] w-96 h-96 bg-violet-400 rounded-full mix-blend-multiply filter blur-[100px] opacity-40 animate-pulse" style={{animationDelay: '2s'}}></div>

      <div className="w-full max-w-[420px] bg-white/80 backdrop-blur-xl rounded-[2.5rem] shadow-soft border border-white p-8 sm:p-10 relative z-10 transition-all">
        
        <div className="flex justify-center mb-8">
          <div className="bg-gradient-to-br from-indigo-600 to-violet-600 p-4 rounded-[1.5rem] shadow-float">
            <BookOpen className="w-8 h-8 text-white" />
          </div>
        </div>
        
        <div className="text-center mb-8">
          <h2 className="text-3xl font-black text-slate-900 tracking-tight mb-2">RTU Portal</h2>
          <p className="text-slate-500 font-medium text-sm flex items-center justify-center gap-1.5">
            {isLogin ? 'Welcome back to your workspace' : 'Create your OJT workspace'} <Sparkles className="w-4 h-4 text-amber-400"/>
          </p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Email Address</label>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
              className="w-full px-5 py-3.5 bg-slate-50/50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all font-semibold text-slate-800 placeholder:text-slate-400 placeholder:font-medium" 
              placeholder="student@rtu.edu.ph" />
          </div>
          <div>
            <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Password</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
              className="w-full px-5 py-3.5 bg-slate-50/50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all font-semibold text-slate-800 placeholder:text-slate-400 placeholder:font-medium" 
              placeholder="••••••••" />
          </div>
          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-indigo-600 text-white font-bold py-4 px-4 rounded-2xl shadow-[0_8px_20px_rgb(0,0,0,0.1)] hover:shadow-float transition-all active:scale-[0.97] disabled:opacity-70 disabled:scale-100 mt-6">
            {loading ? <><Loader2 className="w-5 h-5 animate-spin"/> Authenticating...</> : isLogin ? <><ArrowRight className="w-5 h-5"/> Log In Securely</> : <><UserPlus className="w-5 h-5"/> Create Account</>}
          </button>
        </form>
        
        <div className="mt-8 pt-6 border-t border-slate-100 text-center">
          <button onClick={() => setIsLogin(!isLogin)} className="text-sm font-semibold text-slate-500 hover:text-indigo-600 transition-colors">
            {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Log in'}
          </button>
        </div>
      </div>
    </div>
  )
}