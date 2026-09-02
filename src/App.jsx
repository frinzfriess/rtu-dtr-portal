import { useState, useEffect } from 'react'
import { supabase } from './supabaseClient'
import Auth from './components/Auth'
import Dashboard from './components/Dashboard'

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

  if (loading) return <div className="flex items-center justify-center h-screen bg-slate-50"><div className="animate-pulse flex flex-col items-center"><div className="h-12 w-12 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin"></div><p className="mt-4 text-slate-500 font-medium">Loading workspace...</p></div></div>

  return <div className="min-h-screen bg-slate-50">{!session ? <Auth /> : <Dashboard session={session} />}</div>
}