import { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'
import Auth from './components/Auth'
import Dashboard from './components/Dashboard'
import { Shield, Loader2, Sun, Moon } from 'lucide-react'

export default function App() {
  const [session, setSession] = useState(null)
  const [loading, setLoading] = useState(true)
  const [darkMode, setDarkMode] = useState(true)

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session)
      setLoading(false)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => setSession(session))
    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark')
    } else {
      document.documentElement.classList.remove('dark')
    }
  }, [darkMode])

  if (loading) return (
    <div className="flex items-center justify-center min-h-screen bg-slate-950 text-cyan-400">
      <div className="flex flex-col items-center gap-4">
        <Loader2 className="w-10 h-10 animate-spin text-cyan-400" />
        <p className="font-mono text-xs tracking-widest uppercase animate-pulse">Establishing Secure Uplink...</p>
      </div>
    </div>
  )

  return (
    <div className={`min-h-screen w-full transition-colors duration-300 font-sans ${darkMode ? 'dark bg-slate-950 text-slate-100' : 'bg-slate-50 text-slate-900'}`}>
      {/* Global Theme Toggle Floating Button */}
      <div className="fixed bottom-5 right-5 z-50">
        <button 
          onClick={() => setDarkMode(!darkMode)}
          className="flex items-center gap-2 px-4 py-3 rounded-full shadow-2xl bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black font-mono text-xs transition-all active:scale-95 border border-cyan-300 shadow-glow"
          title="Toggle Theme"
        >
          {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
          <span>{darkMode ? 'LIGHT OP' : 'CYBER DARK'}</span>
        </button>
      </div>

      {!session ? <Auth /> : <Dashboard session={session} darkMode={darkMode} />}
    </div>
  )
}