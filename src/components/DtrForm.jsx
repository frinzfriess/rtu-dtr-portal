import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { Plus, UtensilsCrossed, Clock, CheckCircle2 } from 'lucide-react';
import { calculateHoursRendered, getCurrentTime } from '../utils';

export default function DtrForm({ userId, onEntryAdded }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [timeIn, setTimeIn] = useState('08:00');
  const [timeOut, setTimeOut] = useState('17:00');
  const [remarks, setRemarks] = useState('');
  const [deductLunch, setDeductLunch] = useState(true);
  const [loading, setLoading] = useState(false);
  const [showSuccess, setShowSuccess] = useState(false);

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
      
      // Show success animation briefly
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 2000);
      
    } catch (error) { alert("Failed to add entry. " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-white rounded-[2rem] shadow-soft border border-slate-100 p-5 sm:p-8 transition-all">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
        <div>
          <h3 className="text-xl font-black text-slate-800 flex items-center gap-2">Log Timesheet</h3>
          <p className="text-xs text-slate-500 font-medium mt-1">Record your daily OJT hours below.</p>
        </div>
        
        <label className={`w-full sm:w-auto flex items-center justify-center gap-2 cursor-pointer px-4 py-2.5 rounded-xl border transition-all select-none ${deductLunch ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500 hidden" />
          <UtensilsCrossed className="w-4 h-4"/>
          <span className="text-sm font-bold">Deduct 1h Lunch</span>
          <div className={`w-3 h-3 rounded-full ml-1 border ${deductLunch ? 'bg-indigo-500 border-indigo-600' : 'bg-transparent border-slate-300'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-12 gap-4 lg:gap-5 items-end">
        <div className="lg:col-span-3 w-full">
          <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Date</label>
          <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
            className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500" />
        </div>
        
        <div className="lg:col-span-2 w-full relative">
          <div className="flex justify-between items-end mb-1.5 ml-2 pr-1">
             <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest">Time In</label>
             <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 px-2 py-0.5 rounded-md flex items-center gap-1 active:scale-95 transition-all"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
            className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500" />
        </div>

        <div className="lg:col-span-2 w-full relative">
          <div className="flex justify-between items-end mb-1.5 ml-2 pr-1">
             <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest">Time Out</label>
             <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 px-2 py-0.5 rounded-md flex items-center gap-1 active:scale-95 transition-all"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
            className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500" />
        </div>

        <div className="lg:col-span-5 w-full relative">
          <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Remarks / Tasks</label>
          <div className="flex flex-col sm:flex-row gap-3">
             <input type="text" placeholder="E.g., UI design, testing..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
               className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 focus:border-indigo-500" />
             
             <button type="submit" disabled={loading} 
               className={`w-full sm:w-[140px] flex-shrink-0 flex items-center justify-center gap-2 font-bold py-3.5 rounded-2xl transition-all active:scale-[0.96] ${showSuccess ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/30' : 'bg-slate-900 hover:bg-indigo-600 text-white shadow-lg shadow-slate-900/20 hover:shadow-float'}`}>
               {showSuccess ? <><CheckCircle2 className="w-5 h-5"/> Added</> : <><Plus className="w-5 h-5"/> Submit</>}
             </button>
          </div>
        </div>
      </form>
    </div>
  );
}