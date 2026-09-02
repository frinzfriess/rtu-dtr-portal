import { useState } from 'react'
import { supabase } from '../supabaseClient'
import { BookOpen, LogIn, UserPlus, Sparkles } from 'lucide-react'

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
    <div className="flex justify-center items-center min-h-screen p-4 bg-gradient-to-br from-indigo-50 via-white to-violet-50 overflow-hidden relative">
      {/* Decorative Blobs */}
      <div className="absolute top-[-10%] left-[-10%] w-72 h-72 bg-indigo-300 rounded-full mix-blend-multiply filter blur-3xl opacity-30 animate-blob"></div>
      <div className="absolute top-[20%] right-[-10%] w-72 h-72 bg-violet-300 rounded-full mix-blend-multiply filter blur-3xl opacity-30 animate-blob animation-delay-2000"></div>
      <div className="absolute bottom-[-20%] left-[20%] w-72 h-72 bg-pink-300 rounded-full mix-blend-multiply filter blur-3xl opacity-30 animate-blob animation-delay-4000"></div>

      <div className="w-full max-w-md bg-white/70 backdrop-blur-2xl rounded-[2rem] shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-white p-8 sm:p-10 relative z-10">
        <div className="flex justify-center mb-8">
          <div className="bg-gradient-to-tr from-indigo-600 to-violet-500 p-4 rounded-3xl shadow-lg shadow-indigo-200">
            <BookOpen className="w-8 h-8 text-white" />
          </div>
        </div>
        <h2 className="text-3xl font-black text-center text-slate-900 mb-2 tracking-tight">RTU Portal</h2>
        <p className="text-slate-500 text-center mb-8 font-medium flex items-center justify-center gap-2">
          {isLogin ? 'Welcome back, student!' : 'Create your OJT workspace'} <Sparkles className="w-4 h-4 text-amber-400"/>
        </p>

        <form onSubmit={handleAuth} className="space-y-5">
          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Email Address</label>
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required className="w-full px-5 py-4 bg-white/50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all font-medium text-slate-800 placeholder:text-slate-400 shadow-sm" placeholder="student@rtu.edu.ph" />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Password</label>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-5 py-4 bg-white/50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none transition-all font-medium text-slate-800 placeholder:text-slate-400 shadow-sm" placeholder="••••••••" />
          </div>
          <button type="submit" disabled={loading} className="w-full flex items-center justify-center gap-2 bg-slate-900 hover:bg-indigo-600 text-white font-bold py-4 px-4 rounded-2xl shadow-xl shadow-slate-200 hover:shadow-indigo-200 transition-all active:scale-[0.98] disabled:opacity-70 mt-4">
            {loading ? 'Processing...' : isLogin ? <><LogIn className="w-5 h-5"/> Sign In to Workspace</> : <><UserPlus className="w-5 h-5"/> Create Account</>}
          </button>
        </form>
        <div className="mt-8 text-center">
          <button onClick={() => setIsLogin(!isLogin)} className="text-sm font-bold text-slate-500 hover:text-indigo-600 transition-colors">
            {isLogin ? "First time here? Create an account" : 'Already registered? Sign In'}
          </button>
        </div>
      </div>
    </div>
  )
}