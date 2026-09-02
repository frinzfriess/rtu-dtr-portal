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
      setTimeout(() => setSuccessStatus(false), 2500);
    } catch (error) { alert("Telemetry upload failed: " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-cyber-900 border border-slate-800 rounded-3xl shadow-xl p-6 md:p-8 font-sans relative overflow-hidden">
      <div className="absolute top-0 right-0 w-48 h-48 bg-cyan-500/5 rounded-full blur-3xl pointer-events-none"></div>

      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4 border-b border-slate-800 pb-4">
        <div className="flex items-center gap-3">
          <div className="bg-cyan-500/10 p-2.5 rounded-xl text-cyan-400 border border-cyan-500/20">
            <Terminal className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-base font-black text-white uppercase tracking-wider">Log Telemetry (DTR Entry)</h3>
            <p className="text-xs text-slate-400 font-mono">Record operational hours and security tasks</p>
          </div>
        </div>

        <label className={`flex items-center gap-2 cursor-pointer px-4 py-2.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-300' : 'bg-slate-800 border-slate-700 text-slate-400'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <span className="text-xs font-bold font-mono">Auto-deduct 1h Lunch</span>
          <div className={`w-3 h-3 rounded-full ml-1 border ${deductLunch ? 'bg-cyan-400 border-cyan-300 shadow-[0_0_8px_rgba(6,182,212,0.6)]' : 'bg-transparent border-slate-600'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-5 gap-5 items-end">
        <div>
          <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest mb-2 font-mono">Operation Date</label>
          <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
            className="w-full px-4 py-3.5 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 font-mono" />
        </div>
        <div>
          <div className="flex justify-between items-end mb-2 pr-1">
             <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest font-mono">Time In</label>
             <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[9px] font-bold text-cyan-400 bg-cyan-500/10 px-2 py-0.5 rounded flex items-center gap-1 font-mono"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
            className="w-full px-4 py-3.5 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 font-mono" />
        </div>
        <div>
          <div className="flex justify-between items-end mb-2 pr-1">
             <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest font-mono">Time Out</label>
             <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[9px] font-bold text-cyan-400 bg-cyan-500/10 px-2 py-0.5 rounded flex items-center gap-1 font-mono"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
            className="w-full px-4 py-3.5 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 font-mono" />
        </div>
        <div className="xl:col-span-2 flex flex-col sm:flex-row gap-5 items-end">
          <div className="flex-1 w-full">
            <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest mb-2 font-mono">Incident / Task Remarks</label>
            <input type="text" placeholder="E.g., SOC monitoring, vulnerability assessment..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
              className="w-full px-4 py-3.5 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 placeholder:text-slate-600 font-sans" />
          </div>
          <button type="submit" disabled={loading} 
            className={`w-full sm:w-auto font-bold px-8 py-3.5 rounded-xl transition-all active:scale-95 font-mono tracking-wide h-[48px] whitespace-nowrap ${successStatus ? 'bg-emerald-500 text-slate-950 shadow-neon' : 'bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950 shadow-glow'}`}>
            {successStatus ? <><CheckCircle2 className="w-4 h-4 inline mr-1"/> Logged</> : 'Commit Record'}
          </button>
        </div>
      </form>
    </div>
  );
}