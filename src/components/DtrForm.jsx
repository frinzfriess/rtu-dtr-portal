import { useState } from 'react';
import { supabase } from '../supabaseClient';
import { PlusCircle, UtensilsCrossed, Clock } from 'lucide-react';
import { calculateHoursRendered, getCurrentTime } from '../utils';

export default function DtrForm({ userId, onEntryAdded }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [timeIn, setTimeIn] = useState('08:00');
  const [timeOut, setTimeOut] = useState('17:00');
  const [remarks, setRemarks] = useState('');
  const [deductLunch, setDeductLunch] = useState(true);
  const [loading, setLoading] = useState(false);

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
    } catch (error) { alert("Failed to add entry. " + error.message); } 
    finally { setLoading(false); }
  };

  return (
    <div className="bg-white rounded-[2rem] shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-slate-100 p-5 sm:p-8">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
        <h3 className="text-xl font-black text-slate-800 flex items-center gap-2"><PlusCircle className="w-5 h-5 text-indigo-500"/> Log New Hours</h3>
        <label className={`w-full sm:w-auto flex items-center justify-center space-x-2 cursor-pointer px-4 py-3 rounded-2xl border transition-all ${deductLunch ? 'bg-indigo-50 border-indigo-200 text-indigo-700' : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'}`}>
          <input type="checkbox" checked={deductLunch} onChange={(e) => setDeductLunch(e.target.checked)} className="w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500" />
          <span className="text-sm font-bold flex items-center gap-1.5"><UtensilsCrossed className="w-4 h-4"/> Auto-deduct 1hr lunch</span>
        </label>
      </div>

      <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 items-end">
        <div className="w-full">
          <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Date</label>
          <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500" />
        </div>
        
        {/* Time In with Quick Action */}
        <div className="w-full relative">
          <div className="flex justify-between mb-2 ml-1 pr-1">
             <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider">Time In</label>
             <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500" />
        </div>

        {/* Time Out with Quick Action */}
        <div className="w-full relative">
          <div className="flex justify-between mb-2 ml-1 pr-1">
             <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider">Time Out</label>
             <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1"><Clock className="w-3 h-3"/> Now</button>
          </div>
          <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500" />
        </div>

        <div className="w-full">
          <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 ml-1">Remarks / Tasks</label>
          <input type="text" placeholder="E.g., UI design, bug fixing..." value={remarks} onChange={(e) => setRemarks(e.target.value)} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500" />
        </div>

        <div className="lg:col-span-4 mt-2">
          <button type="submit" disabled={loading} className="w-full bg-slate-900 hover:bg-indigo-600 text-white font-bold py-4 rounded-2xl shadow-lg shadow-slate-200 hover:shadow-indigo-200 transition-all active:scale-[0.98] disabled:opacity-50">
            {loading ? 'Saving Record...' : 'Submit Daily Record'}
          </button>
        </div>
      </form>
    </div>
  );
}