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
      setTimeout(() => setSuccessStatus(false), 2500);
    } catch (error) { alert("Failed to add record: " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-dark-900 border border-slate-800 rounded-3xl shadow-soft p-6 md:p-8 w-full relative overflow-hidden">
      <div className="absolute top-0 right-0 w-48 h-48 bg-blue-500/5 rounded-full blur-3xl pointer-events-none"></div>

      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4 border-b border-slate-800 pb-4">
        <div className="flex items-center gap-3">
          <div className="bg-blue-500/10 p-2.5 rounded-xl text-blue-400 border border-blue-500/20">
            <Clock className="w-5 h-5" />
          </div>
          <div>
            <h3 className="text-base font-black text-white">Log Daily Time Record</h3>
            <p className="text-xs text-slate-400">Record your working hours and daily tasks</p>
          </div>
        </div>

        <label className={`flex items-center gap-2 cursor-pointer px-4 py-2.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-blue-500/10 border-blue-500/30 text-blue-300' : 'bg-dark-950 border-slate-800 text-slate-400'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <span className="text-xs font-bold">Auto-deduct 1h Lunch</span>
          <div className={`w-3 h-3 rounded-full ml-1 border ${deductLunch ? 'bg-blue-500 border-blue-400 shadow-[0_0_8px_rgba(59,130,246,0.5)]' : 'bg-transparent border-slate-600'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-5 gap-5 items-end">
        <div>
          <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2">Date</label>
          <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
            className="w-full px-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20" />
        </div>
        <div>
          <div className="flex justify-between items-end mb-2 pr-1">
             <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider">Time In</label>
             <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-blue-400 bg-blue-500/10 hover:bg-blue-500/20 px-2 py-0.5 rounded flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
            className="w-full px-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20" />
        </div>
        <div>
          <div className="flex justify-between items-end mb-2 pr-1">
             <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider">Time Out</label>
             <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-blue-400 bg-blue-500/10 hover:bg-blue-500/20 px-2 py-0.5 rounded flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
            className="w-full px-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20" />
        </div>
        <div className="xl:col-span-2 flex flex-col sm:flex-row gap-5 items-end">
          <div className="flex-1 w-full">
            <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-2">Remarks / Tasks</label>
            <input type="text" placeholder="E.g., System updates, coding, testing..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
              className="w-full px-4 py-3.5 bg-dark-950 border border-slate-800 rounded-xl text-sm font-bold text-white outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-500/20 placeholder:text-slate-600" />
          </div>
          <button type="submit" disabled={loading} 
            className={`w-full sm:w-auto font-bold px-8 py-3.5 rounded-xl transition-all active:scale-95 h-[48px] whitespace-nowrap shadow-lg ${successStatus ? 'bg-emerald-500 text-white shadow-emerald-500/20' : 'bg-blue-600 hover:bg-blue-500 text-white shadow-blue-600/25'}`}>
            {successStatus ? <><CheckCircle2 className="w-4 h-4 inline mr-1"/> Logged</> : 'Add Record'}
          </button>
        </div>
      </form>
    </div>
  );
}