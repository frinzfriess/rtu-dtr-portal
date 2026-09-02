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
      setShowSuccess(true);
      setTimeout(() => setShowSuccess(false), 2000);
    } catch (error) { alert("Failed to add entry. " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-white rounded-2xl shadow-soft border border-slate-100 p-5 sm:p-6 w-full">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-5 gap-3">
        <div>
          <h3 className="text-lg font-black text-slate-900 flex items-center gap-2">Add Time Record</h3>
        </div>
        <label className={`w-full sm:w-auto flex items-center justify-center gap-2 cursor-pointer px-3 py-2 rounded-lg border transition-all select-none ${deductLunch ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="hidden" />
          <UtensilsCrossed className="w-3.5 h-3.5"/>
          <span className="text-xs font-bold">Auto-deduct 1h Lunch</span>
          <div className={`w-2.5 h-2.5 rounded-full ml-1 border ${deductLunch ? 'bg-indigo-500 border-indigo-600' : 'bg-transparent border-slate-300'}`}></div>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="space-y-4">
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="w-full">
            <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1">Date</label>
            <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
              className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500" />
          </div>
          
          <div className="grid grid-cols-2 gap-4">
            <div className="w-full relative">
              <div className="flex justify-between items-end mb-1.5 ml-1 pr-1">
                 <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide truncate">Time In</label>
                 <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 px-1.5 py-0.5 rounded flex items-center gap-0.5 active:scale-95 flex-shrink-0"><Clock className="w-2.5 h-2.5"/> Now</button>
              </div>
              <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
                className="w-full px-2 sm:px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500" />
            </div>

            <div className="w-full relative">
              <div className="flex justify-between items-end mb-1.5 ml-1 pr-1">
                 <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide truncate">Time Out</label>
                 <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-indigo-600 hover:text-indigo-800 bg-indigo-50 px-1.5 py-0.5 rounded flex items-center gap-0.5 active:scale-95 flex-shrink-0"><Clock className="w-2.5 h-2.5"/> Now</button>
              </div>
              <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
                className="w-full px-2 sm:px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500" />
            </div>
          </div>
        </div>

        <div className="w-full">
          <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1">Remarks / Tasks</label>
          <input type="text" placeholder="E.g., UI design, fixing bugs..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
            className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-700 outline-none transition-all focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500" />
        </div>

        <button type="submit" disabled={loading} 
          className={`w-full flex items-center justify-center gap-2 font-bold py-3 rounded-xl transition-all active:scale-[0.98] ${showSuccess ? 'bg-emerald-500 text-white shadow-md shadow-emerald-500/20' : 'bg-slate-900 hover:bg-indigo-600 text-white shadow-md shadow-slate-900/10'}`}>
          {showSuccess ? <><CheckCircle2 className="w-4 h-4"/> Success</> : <><Plus className="w-4 h-4"/> Submit Record</>}
        </button>
      </form>
    </div>
  );
}