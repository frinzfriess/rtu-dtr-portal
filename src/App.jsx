import { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'
import Auth from './components/Auth'
import Dashboard from './components/Dashboard'
import { Loader2 } from 'lucide-react'

export default function App() {
  const [session, setSession] = useState(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session)
      setLoading(false)
    })
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => setSession(session))
    return () => subscription.unsubscribe()
  }, [])

  if (loading) return (
    <div className="flex items-center justify-center min-h-screen bg-cyber-950 text-cyan-400">
      <div className="flex flex-col items-center gap-4">
        <Loader2 className="w-10 h-10 animate-spin text-cyan-400" />
        <p className="font-mono text-xs tracking-widest uppercase animate-pulse">Establishing Secure Uplink...</p>
      </div>
    </div>
  )

  return <div className="min-h-screen w-full bg-cyber-950 font-sans text-slate-100">{!session ? <Auth /> : <Dashboard session={session} />}</div>
}