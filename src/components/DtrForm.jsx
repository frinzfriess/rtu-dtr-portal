import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { PlusCircle, UtensilsCrossed, Clock, CheckCircle2, Terminal } from 'lucide-react';
import { calculateHoursRendered, getCurrentTime } from '../utils';

export default function DtrForm({ userId, onEntryAdded }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [timeIn, setTimeIn] = useState('08:00');
  const [timeOut, setTimeOut] = useState('17:00');
  const [remarks, setRemarks] = useState('');
  const [deductLunch, setDeductLunch] = useState(true);
  const [loading, setLoading] = useState(false);
  const [successStatus, setSuccessStatus] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    const hours_rendered = calculateHoursRendered(timeIn, timeOut, deductLunch);

    if (hours_rendered <= 0) {
      alert("Invalid time range. Time out must be after Time in.");
      setLoading(false);
      return;
    }

    try {
      const { error } = await supabase.from('dtr_entries').insert([{
        user_id: userId, date, time_in: `${timeIn}:00`, time_out: `${timeOut}:00`, hours_rendered, remarks
      }]);
      if (error) throw error;
      
      setRemarks('');
      onEntryAdded();
      setSuccessStatus(true);
      setTimeout(() => setSuccessStatus(false), 2000);
    } catch (error) { alert("Failed to add entry. " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl shadow-soft p-5 sm:p-6 md:p-8 w-full transition-colors mb-8 min-w-0">
      
      {/* Header & Checkbox row with bulletproof flex wrapping for mobile */}
      <div className="flex flex-col gap-4 mb-6 pb-4 border-b border-slate-100 dark:border-slate-800 min-w-0">
        <div className="flex items-center gap-3 min-w-0">
          <div className="bg-cyan-500/10 p-2.5 rounded-xl text-cyan-600 dark:text-cyan-400 border border-cyan-500/30 flex-shrink-0 shadow-glow">
            <Terminal className="w-5 h-5" />
          </div>
          <div className="min-w-0">
            <h3 className="text-base font-black text-slate-900 dark:text-white uppercase tracking-wider font-mono truncate">Log New Hours</h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 font-mono truncate">Record your daily OJT hours below.</p>
          </div>
        </div>

        <div className="flex items-center justify-start min-w-0">
          <label className={`inline-flex items-center gap-2 cursor-pointer px-4 py-2.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-700 dark:text-cyan-300' : 'bg-slate-50 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-500'}`}>
            <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
            <UtensilsCrossed className="w-4 h-4 flex-shrink-0"/>
            <span className="text-xs font-bold font-mono">Auto-deduct 1h lunch</span>
            <div className={`w-3 h-3 rounded-full ml-1 border flex-shrink-0 ${deductLunch ? 'bg-cyan-500 border-cyan-600 dark:bg-cyan-400 dark:border-cyan-300' : 'bg-transparent border-slate-400'}`}></div>
          </label>
        </div>
      </div>

      <form onSubmit={handleSubmit} className="space-y-4 min-w-0">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 sm:gap-5 min-w-0">
          <div className="w-full min-w-0">
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-2 font-mono">Date</label>
            <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all min-w-0 font-mono" />
          </div>
          <div className="w-full min-w-0">
            <div className="flex justify-between items-end mb-2 pr-1 min-w-0">
               <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest truncate font-mono">Time In</label>
               <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2.5 py-1 rounded flex items-center gap-1 flex-shrink-0 font-mono"><Clock className="w-3 h-3"/> Now</button>
            </div>
            <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all min-w-0 font-mono" />
          </div>
          <div className="w-full min-w-0">
            <div className="flex justify-between items-end mb-2 pr-1 min-w-0">
               <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest truncate font-mono">Time Out</label>
               <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2.5 py-1 rounded flex items-center gap-1 flex-shrink-0 font-mono"><Clock className="w-3 h-3"/> Now</button>
            </div>
            <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all min-w-0 font-mono" />
          </div>
        </div>

        <div className="w-full min-w-0">
          <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-2 font-mono">Remarks / Tasks</label>
          <input type="text" placeholder="E.g., UI design, system maintenance, coding..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
            className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all placeholder:text-slate-400 dark:placeholder:text-slate-600 min-w-0 font-sans" />
        </div>

        <div className="flex justify-end pt-2">
          <button type="submit" disabled={loading} 
            className={`w-full sm:w-auto font-bold px-8 py-3.5 rounded-xl transition-all active:scale-95 shadow-glow font-mono text-xs tracking-wide flex-shrink-0 ${successStatus ? 'bg-emerald-500 text-slate-950 shadow-neon' : 'bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950'}`}>
            {successStatus ? <><CheckCircle2 className="w-4 h-4 inline mr-1"/> Logged</> : 'Submit Daily Record'}
          </button>
        </div>
      </form>
    </div>
  );
}