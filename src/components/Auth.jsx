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
    <div className="flex justify-center items-center min-h-screen p-4 sm:p-8 bg-slate-50 relative overflow-hidden">
      {/* Decorative Background Elements */}
      <div className="absolute top-[-10%] left-[-5%] w-72 h-72 bg-indigo-300 rounded-full mix-blend-multiply filter blur-[80px] opacity-40"></div>
      <div className="absolute bottom-[-10%] right-[-5%] w-72 h-72 bg-violet-300 rounded-full mix-blend-multiply filter blur-[80px] opacity-40"></div>

      <div className="w-full max-w-[400px] bg-white rounded-3xl shadow-soft border border-slate-100 p-8 relative z-10">
        <div className="flex justify-center mb-6">
          <div className="bg-indigo-600 p-3.5 rounded-2xl shadow-float">
            <BookOpen className="w-7 h-7 text-white" />
          </div>
        </div>
        
        <div className="text-center mb-8">
          <h2 className="text-2xl font-black text-slate-900 tracking-tight mb-1.5">RTU DTR Portal</h2>
          <p className="text-slate-500 font-medium text-sm flex items-center justify-center gap-1.5">
            {isLogin ? 'Sign in to your workspace' : 'Create your workspace'} <Sparkles className="w-4 h-4 text-amber-400"/>
          </p>
        </div>

        <form onSubmit={handleAuth} className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1.5 ml-1">RTU Email</label>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required 
              className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all font-semibold text-slate-800 placeholder:text-slate-400" 
              placeholder="student@rtu.edu.ph" />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1.5 ml-1">Password</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required 
              className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all font-semibold text-slate-800 placeholder:text-slate-400" 
              placeholder="••••••••" />
          </div>
          <button type="submit" disabled={loading} 
            className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-indigo-600 text-white font-bold py-3.5 px-4 rounded-xl shadow-lg shadow-slate-900/10 hover:shadow-indigo-500/25 transition-all active:scale-[0.98] disabled:opacity-70 mt-6">
            {loading ? <><Loader2 className="w-4 h-4 animate-spin"/> Processing...</> : isLogin ? <><ArrowRight className="w-4 h-4"/> Sign In</> : <><UserPlus className="w-4 h-4"/> Register</>}
          </button>
        </form>
        
        <div className="mt-6 pt-5 border-t border-slate-100 text-center">
          <button onClick={() => setIsLogin(!isLogin)} className="text-sm font-semibold text-indigo-600 hover:text-indigo-800 transition-colors">
            {isLogin ? "Don't have an account? Sign up" : 'Already have an account? Log in'}
          </button>
        </div>
      </div>
    </div>
  )
}