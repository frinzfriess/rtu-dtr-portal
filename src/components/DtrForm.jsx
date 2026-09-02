import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { PlusCircle, UtensilsCrossed, Clock, CheckCircle2 } from 'lucide-react';
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
    <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl shadow-soft p-6 md:p-8 w-full transition-colors mb-8">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4 border-b border-slate-100 dark:border-slate-800 pb-4">
        <div className="flex items-center gap-2.5">
          <div className="bg-cyan-500/10 p-2.5 rounded-xl text-cyan-600 dark:text-cyan-400 border border-cyan-500/20">
            <PlusCircle className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-base font-black text-slate-900 dark:text-white">Log New Hours</h3>
            <p className="text-xs text-slate-500 dark:text-slate-400">Record your daily OJT hours below.</p>
          </div>
        </div>

        <label className={`flex items-center gap-2 cursor-pointer px-4 py-2.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-700 dark:text-cyan-300' : 'bg-slate-50 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-500'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <UtensilsCrossed className="w-4 h-4"/>
          <span className="text-xs font-bold">Auto-deduct 1h lunch</span>
          <div className={`w-3 h-3 rounded-full ml-1 border ${deductLunch ? 'bg-cyan-500 border-cyan-600 dark:bg-cyan-400 dark:border-cyan-300' : 'bg-transparent border-slate-400'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="space-y-5">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div>
            <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-2">Date</label>
            <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all" />
          </div>
          <div>
            <div className="flex justify-between items-end mb-2 pr-1">
               <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Time In</label>
               <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2 py-0.5 rounded flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
            </div>
            <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all" />
          </div>
          <div>
            <div className="flex justify-between items-end mb-2 pr-1">
               <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Time Out</label>
               <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2 py-0.5 rounded flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
            </div>
            <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
              className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all" />
          </div>
        </div>

        <div>
          <label className="block text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-2">Remarks / Tasks</label>
          <input type="text" placeholder="E.g., UI design, bug fixing..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
            className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all placeholder:text-slate-400 dark:placeholder:text-slate-600" />
        </div>

        <div className="flex justify-end pt-2">
          <button type="submit" disabled={loading} 
            className={`w-full sm:w-auto font-bold px-8 py-3.5 rounded-xl transition-all active:scale-95 shadow-lg ${successStatus ? 'bg-emerald-500 text-white shadow-emerald-500/20' : 'bg-slate-900 dark:bg-cyan-500 hover:bg-cyan-600 dark:hover:bg-cyan-400 text-white dark:text-slate-950 shadow-slate-900/15'}`}>
            {successStatus ? <><CheckCircle2 className="w-4 h-4 inline mr-1"/> Logged</> : 'Submit Daily Record'}
          </button>
        </div>
      </form>
    </div>
  );
}