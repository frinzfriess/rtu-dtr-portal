import { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'
import Auth from './components/Auth'
import Dashboard from './components/Dashboard'
import { Loader2 } from 'lucide-react'

export default function App() {
  const [session, setSession] = useState(null)
  const [loading, setLoading] = useState(true)
  const [darkMode, setDarkMode] = useState(() => {
    const saved = localStorage.getItem('rtu_dtr_theme');
    return saved ? JSON.parse(saved) : false;
  })

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session)
      setLoading(false)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => setSession(session))
    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    localStorage.setItem('rtu_dtr_theme', JSON.stringify(darkMode));
    if (darkMode) {
      document.documentElement.classList.add('dark')
    } else {
      document.documentElement.classList.remove('dark')
    }
  }, [darkMode])

  if (loading) return (
    <div className="flex items-center justify-center min-h-screen bg-slate-100 dark:bg-slate-950 text-cyan-600 dark:text-cyan-400">
      <div className="flex flex-col items-center gap-4">
        <Loader2 className="w-10 h-10 animate-spin text-cyan-600 dark:text-cyan-400" />
        <p className="font-mono text-xs tracking-widest uppercase animate-pulse">Initializing Portal...</p>
      </div>
    </div>
  )

  return (
    <div className={`min-h-screen w-full transition-colors duration-300 font-sans ${darkMode ? 'dark bg-slate-950 text-slate-100' : 'bg-slate-100 text-slate-900'}`}>
      {!session ? <Auth /> : <Dashboard session={session} darkMode={darkMode} setDarkMode={setDarkMode} />}
    </div>
  )
}