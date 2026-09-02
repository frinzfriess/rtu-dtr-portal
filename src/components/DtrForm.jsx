import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { Shield, Clock, PlusCircle, CheckCircle2, Terminal } from 'lucide-react';
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
    <div className="bg-cyber-900 border border-slate-800 rounded-2xl shadow-xl p-5 sm:p-6 w-full font-sans relative overflow-hidden">
      <div className="absolute top-0 right-0 w-32 h-32 bg-cyan-500/5 rounded-full blur-2xl pointer-events-none"></div>

      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-5 gap-3 border-b border-slate-800 pb-4">
        <div className="flex items-center gap-2.5">
          <div className="bg-cyan-500/10 p-2 rounded-lg text-cyan-400 border border-cyan-500/20">
            <Terminal className="w-4 h-4" />
          </div>
          <div>
            <h3 className="text-sm font-black text-white uppercase tracking-wider">Log Telemetry (DTR Entry)</h3>
            <p className="text-[10px] text-slate-400 font-mono">Record operational hours and cybersecurity tasks</p>
          </div>
        </div>

        <label className={`w-full sm:w-auto flex items-center justify-center gap-2 cursor-pointer px-3.5 py-2 rounded-xl border transition-all select-none ${deductLunch ? 'bg-cyan-500/10 border-cyan-500/30 text-cyan-300' : 'bg-slate-800 border-slate-700 text-slate-400'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <span className="text-xs font-bold font-mono">Auto-deduct 1h Lunch</span>
          <div className={`w-2.5 h-2.5 rounded-full ml-1 border ${deductLunch ? 'bg-cyan-400 border-cyan-300 shadow-[0_0_8px_rgba(6,182,212,0.6)]' : 'bg-transparent border-slate-600'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="space-y-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="w-full">
            <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest mb-1.5 ml-1 font-mono">Operation Date</label>
            <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
              className="w-full px-3.5 py-3 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 transition-all font-mono" />
          </div>
          
          <div className="grid grid-cols-2 gap-3">
            <div className="w-full">
              <div className="flex justify-between items-end mb-1.5 ml-1 pr-1">
                 <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest font-mono">Time In</label>
                 <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[9px] font-bold text-cyan-400 bg-cyan-500/10 hover:bg-cyan-500/20 px-1.5 py-0.5 rounded flex items-center gap-0.5 font-mono"><Clock className="w-2.5 h-2.5"/> Now</button>
              </div>
              <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
                className="w-full px-2 sm:px-3 py-3 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 transition-all font-mono" />
            </div>

            <div className="w-full">
              <div className="flex justify-between items-end mb-1.5 ml-1 pr-1">
                 <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest font-mono">Time Out</label>
                 <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[9px] font-bold text-cyan-400 bg-cyan-500/10 hover:bg-cyan-500/20 px-1.5 py-0.5 rounded flex items-center gap-0.5 font-mono"><Clock className="w-2.5 h-2.5"/> Now</button>
              </div>
              <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
                className="w-full px-2 sm:px-3 py-3 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 transition-all font-mono" />
            </div>
          </div>
        </div>

        <div className="w-full">
          <label className="block text-[10px] font-bold text-cyan-400 uppercase tracking-widest mb-1.5 ml-1 font-mono">Incident / Task Remarks</label>
          <input type="text" placeholder="E.g., Vulnerability assessment, SOC monitoring, firewall config..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
            className="w-full px-3.5 py-3 bg-cyber-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 transition-all placeholder:text-slate-600" />
        </div>

        <button type="submit" disabled={loading} 
          className={`w-full flex items-center justify-center gap-2 font-bold py-3.5 rounded-xl transition-all active:scale-[0.98] font-mono tracking-wide ${successStatus ? 'bg-emerald-500 text-slate-950 shadow-neon' : 'bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 hover:to-indigo-500 text-slate-950 shadow-glow'}`}>
          {successStatus ? <><CheckCircle2 className="w-4 h-4"/> Telemetry Logged Successfully</> : <><PlusCircle className="w-4 h-4"/> Commit Telemetry Record</>}
        </button>
      </form>
    </div>
  );
}