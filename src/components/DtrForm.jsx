import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { Clock, PlusCircle, CheckCircle2, Terminal, Shield } from 'lucide-react';
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
      alert("Invalid time vector. Time out must succeed time in.");
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
    } catch (error) { alert("Telemetry upload failed: " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl shadow-sm p-4 sm:p-6 w-full font-sans relative overflow-hidden transition-colors">
      <div className="absolute top-0 right-0 w-24 h-24 bg-cyan-500/5 rounded-full blur-xl pointer-events-none"></div>

      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4 gap-3 border-b border-slate-100 dark:border-slate-800 pb-3">
        <div className="flex items-center gap-2">
          <div className="bg-cyan-500/10 p-2 rounded-lg text-cyan-600 dark:text-cyan-400 border border-cyan-500/20">
            <Terminal className="w-4 h-4" />
          </div>
          <div>
            <h3 className="text-xs font-black text-slate-900 dark:text-white uppercase tracking-wider font-mono">Log Telemetry (DTR)</h3>
            <p className="text-[10px] text-slate-500 dark:text-slate-400 font-mono">Record operational hours & incident notes</p>
          </div>
        </div>

        <label className={`w-full sm:w-auto flex items-center justify-center gap-2 cursor-pointer px-3 py-1.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-700 dark:text-cyan-300' : 'bg-slate-100 dark:bg-slate-800 border-slate-200 dark:border-slate-700 text-slate-500'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <span className="text-[11px] font-bold font-mono">Auto-deduct 1h Lunch</span>
          <div className={`w-2.5 h-2.5 rounded-full ml-1 border ${deductLunch ? 'bg-cyan-500 border-cyan-600 dark:bg-cyan-400 dark:border-cyan-300' : 'bg-transparent border-slate-400'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="space-y-3.5">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div className="w-full">
            <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-0.5 font-mono">Operation Date</label>
            <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
              className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all font-mono" />
          </div>
          
          <div className="w-full">
            <div className="flex justify-between items-end mb-1 ml-0.5 pr-1">
               <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time In</label>
               <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[9px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-1.5 py-0.5 rounded flex items-center gap-0.5 font-mono"><Clock className="w-2.5 h-2.5"/> Now</button>
            </div>
            <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
              className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all font-mono" />
          </div>

          <div className="w-full">
            <div className="flex justify-between items-end mb-1 ml-0.5 pr-1">
               <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time Out</label>
               <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[9px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-1.5 py-0.5 rounded flex items-center gap-0.5 font-mono"><Clock className="w-2.5 h-2.5"/> Now</button>
            </div>
            <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
              className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all font-mono" />
          </div>
        </div>

        <div className="w-full">
          <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-0.5 font-mono">Incident / Task Remarks</label>
          <input type="text" placeholder="E.g., SOC monitoring, vulnerability assessment, firewall config..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
            className="w-full px-3.5 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 transition-all placeholder:text-slate-400 dark:placeholder:text-slate-600 font-sans" />
        </div>

        <button type="submit" disabled={loading} 
          className={`w-full flex items-center justify-center gap-2 font-bold py-3 rounded-xl transition-all active:scale-[0.98] font-mono text-xs tracking-wide ${successStatus ? 'bg-emerald-500 text-slate-950 shadow-neon' : 'bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950 shadow-glow'}`}>
          {successStatus ? <><CheckCircle2 className="w-4 h-4"/> Telemetry Logged</> : <><PlusCircle className="w-4 h-4"/> Commit Telemetry Record</>}
        </button>
      </form>
    </div>
  );
}