import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import DtrForm from './DtrForm';
import { LogOut, Download, Clock, CalendarDays, Target, Building2, GraduationCap, User, Trash2, Edit, X, Save, FileSpreadsheet, Search, SearchX, CheckCircle2 } from 'lucide-react';
import { calculateHoursRendered, getGreeting } from '../utils';

export default function Dashboard({ session }) {
  const [profile, setProfile] = useState(null);
  const [dtr, setDtr] = useState([]);
  const [totalHours, setTotalHours] = useState(0);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');

  // Edit State
  const [editingEntry, setEditingEntry] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [isSavingEdit, setIsSavingEdit] = useState(false);
  const [deductLunchEdit, setDeductLunchEdit] = useState(true);

  useEffect(() => { fetchProfile(); fetchDTR(); }, [session]);

  async function fetchProfile() {
    try {
      const { data, error } = await supabase.from('profiles').select('*').eq('id', session.user.id).single();
      if (error && error.code === 'PGRST116') {
        const { data: newProfile, error: upsertError } = await supabase.from('profiles').upsert([{ 
            id: session.user.id, school: 'Rizal Technological University', course: 'BS Information Technology', 
            company_name: 'Pending Company Assignment', required_hours: 300, full_name: '' 
        }]).select().single();
        if (upsertError) throw upsertError;
        setProfile(newProfile);
      } else { setProfile(data); }
    } catch (error) { console.error(error); }
  }

  async function fetchDTR() {
    try {
      const { data, error } = await supabase.from('dtr_entries').select('*').eq('user_id', session.user.id).order('date', { ascending: false });
      if (error) throw error;
      if (data) {
        setDtr(data);
        setTotalHours(data.reduce((sum, entry) => sum + Number(entry.hours_rendered), 0));
      }
    } catch (error) { console.error(error); } finally { setLoading(false); }
  }

  async function updateProfileInfo(field, value) {
    if (!profile || profile[field] == value) return;
    try {
      await supabase.from('profiles').update({ [field]: value }).eq('id', session.user.id);
      setProfile(prev => ({...prev, [field]: value}));
    } catch (error) { alert('Failed to update profile'); }
  }

  async function deleteEntry(id) {
    if(!window.confirm("Are you sure you want to delete this record? This action cannot be undone.")) return;
    try {
      await supabase.from('dtr_entries').delete().eq('id', id);
      fetchDTR();
    } catch (error) { alert("Could not delete record."); }
  }

  const handleEditClick = (entry) => {
    setEditingEntry(entry);
    setEditForm({
      date: entry.date,
      time_in: entry.time_in.substring(0, 5),
      time_out: entry.time_out.substring(0, 5),
      remarks: entry.remarks || ''
    });
    setDeductLunchEdit(entry.hours_rendered < 5 ? false : true); 
  };

  const saveEdit = async () => {
    setIsSavingEdit(true);
    const newHours = calculateHoursRendered(editForm.time_in, editForm.time_out, deductLunchEdit);
    if (newHours <= 0) {
      alert("Invalid time range.");
      setIsSavingEdit(false); return;
    }
    try {
      const { error } = await supabase.from('dtr_entries').update({
        date: editForm.date, time_in: `${editForm.time_in}:00`, time_out: `${editForm.time_out}:00`,
        hours_rendered: newHours, remarks: editForm.remarks
      }).eq('id', editingEntry.id);
      if (error) throw error;
      setEditingEntry(null); fetchDTR();
    } catch (err) { alert("Failed to update entry."); } 
    finally { setIsSavingEdit(false); }
  };

  const exportStyledExcel = async () => {
    try {
      const ExcelJS = (await import('exceljs')).default || await import('exceljs');
      const { saveAs } = (await import('file-saver')).default || await import('file-saver');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Daily Time Record');
      worksheet.columns = [ { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 40 } ];

      worksheet.mergeCells('A1:E1');
      const title = worksheet.getCell('A1'); title.value = 'RIZAL TECHNOLOGICAL UNIVERSITY'; title.font = { bold: true, size: 14 }; title.alignment = { horizontal: 'center' };
      worksheet.mergeCells('A2:E2');
      const subTitle = worksheet.getCell('A2'); subTitle.value = 'DAILY TIME RECORD (OJT)'; subTitle.font = { bold: true, size: 12 }; subTitle.alignment = { horizontal: 'center' };
      worksheet.addRow([]);
      worksheet.addRow(['Name:', profile.full_name || 'N/A', '', 'Company:', profile.company_name || 'N/A']);
      worksheet.addRow(['Course:', profile.course || 'N/A', '', 'School:', profile.school || 'N/A']);
      worksheet.addRow([]);
      const headerRow = worksheet.addRow(['Date', 'Time In', 'Time Out', 'Hours', 'Remarks/Tasks']);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }; headerRow.alignment = { horizontal: 'center' };
      headerRow.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });

      const sortedDtr = [...dtr].sort((a,b) => new Date(a.date) - new Date(b.date));
      sortedDtr.forEach(entry => {
          const row = worksheet.addRow([ new Date(entry.date).toLocaleDateString(), entry.time_in.substring(0, 5), entry.time_out.substring(0, 5), entry.hours_rendered, entry.remarks || '' ]);
          row.eachCell(cell => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' }; });
      });
      worksheet.addRow([]);
      const totalRow = worksheet.addRow(['', '', 'TOTAL HOURS:', totalHours.toFixed(2), '']);
      totalRow.font = { bold: true, size: 12 }; totalRow.getCell(4).font = { bold: true, size: 12, color: { argb: 'FFDC2626' } };
      
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `DTR_${profile.full_name || 'Student'}.xlsx`);
    } catch (err) { alert("Failed to export Excel. Please make sure dependencies are installed."); }
  };

  if (loading || !profile) return null;
  const progressPercentage = Math.min((totalHours / profile.required_hours) * 100, 100);
  const remainingHours = Math.max(profile.required_hours - totalHours, 0);
  const filteredDtr = dtr.filter(e => e.remarks?.toLowerCase().includes(searchTerm.toLowerCase()) || e.date.includes(searchTerm));

  return (
    <div className="max-w-7xl mx-auto py-6 px-4 sm:px-6 lg:px-8 relative pb-20">
      
      {/* Edit Modal (Mobile Responsive) */}
      {editingEntry && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4 z-50 transition-all">
          <div className="bg-white rounded-t-3xl sm:rounded-3xl p-6 sm:p-8 w-full max-w-md shadow-2xl relative border border-slate-100 animate-slide-up sm:animate-none">
            <button onClick={() => setEditingEntry(null)} className="absolute top-6 right-6 text-slate-400 hover:text-slate-600 bg-slate-100 p-2 rounded-full transition-colors"><X className="w-5 h-5" /></button>
            <h3 className="text-2xl font-black text-slate-800 mb-6 flex items-center gap-2"><Edit className="w-6 h-6 text-indigo-500"/> Edit Record</h3>
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-500 uppercase mb-1 ml-1">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 font-bold outline-none" />
              </div>
              <div className="flex gap-4">
                <div className="flex-1">
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1 ml-1">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 font-bold outline-none" />
                </div>
                <div className="flex-1">
                  <label className="block text-xs font-bold text-slate-500 uppercase mb-1 ml-1">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 font-bold outline-none" />
                </div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-500 uppercase mb-1 ml-1">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-4 focus:ring-indigo-500/10 font-bold outline-none" />
              </div>
              <label className="flex items-center space-x-2 cursor-pointer mt-2 bg-slate-50 p-3 rounded-xl border border-slate-200">
                <input type="checkbox" checked={deductLunchEdit} onChange={(e) => setDeductLunchEdit(e.target.checked)} className="w-5 h-5 text-indigo-600 rounded border-slate-300" />
                <span className="text-sm font-bold text-slate-600">Auto-deduct 1hr lunch break</span>
              </label>
            </div>
            <div className="flex gap-3 mt-8">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-4 rounded-2xl transition-colors">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-indigo-200 transition-all flex items-center justify-center gap-2">
                <Save className="w-5 h-5"/> {isSavingEdit ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Top Header / Sticky Mobile Nav */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4 bg-slate-50/80 backdrop-blur-xl sticky top-0 z-30 py-4 border-b md:border-none border-slate-200/50 -mx-4 px-4 md:mx-0 md:px-0">
        <div>
          <p className="text-indigo-600 font-bold text-sm tracking-wide uppercase mb-1">{getGreeting()}</p>
          <h1 className="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight leading-none">{profile.full_name || session.user.email.split('@')[0]}</h1>
        </div>
        <div className="flex gap-2 w-full md:w-auto">
          <button onClick={exportStyledExcel} className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-emerald-600 text-white hover:bg-emerald-700 shadow-lg shadow-emerald-200/50 font-bold px-4 py-3 rounded-2xl transition-all active:scale-95 text-sm sm:text-base">
            <FileSpreadsheet className="w-4 h-4" /> Export DTR
          </button>
          <button onClick={() => supabase.auth.signOut()} className="flex-none flex items-center justify-center bg-rose-50 text-rose-600 hover:bg-rose-100 font-bold p-3 rounded-2xl border border-rose-100 transition-all active:scale-95">
            <LogOut className="w-5 h-5" />
          </button>
        </div>
      </div>
      
      {/* Profile & Stats Layout */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
        {/* Left Side: Profile Cards */}
        <div className="lg:col-span-2 bg-white p-5 sm:p-8 rounded-[2rem] shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-slate-100">
          <h2 className="text-lg font-black text-slate-800 mb-6 flex items-center gap-2"><User className="w-5 h-5 text-indigo-500"/> Student Identity</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 focus-within:ring-4 focus-within:ring-indigo-500/10 focus-within:border-indigo-300 transition-all">
              <label className="text-xs uppercase tracking-wider text-slate-500 font-bold flex items-center gap-1.5 mb-2"><User className="w-3 h-3"/> Full Name</label>
              <input type="text" defaultValue={profile.full_name || ''} placeholder="Juan Dela Cruz" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 text-lg" />
            </div>
            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 focus-within:ring-4 focus-within:ring-indigo-500/10 focus-within:border-indigo-300 transition-all">
              <label className="text-xs uppercase tracking-wider text-slate-500 font-bold flex items-center gap-1.5 mb-2"><Building2 className="w-3 h-3"/> Company</label>
              <input type="text" defaultValue={profile.company_name} placeholder="Tech Corp Inc." onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 text-lg" />
            </div>
            <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 focus-within:ring-4 focus-within:ring-indigo-500/10 focus-within:border-indigo-300 transition-all sm:col-span-2">
              <label className="text-xs uppercase tracking-wider text-slate-500 font-bold flex items-center gap-1.5 mb-2"><GraduationCap className="w-3 h-3"/> School / University</label>
              <input type="text" defaultValue={profile.school} onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 text-lg" />
            </div>
          </div>
        </div>

        {/* Right Side: Awesome Stats Card */}
        <div className="bg-gradient-to-br from-indigo-600 via-violet-600 to-purple-700 p-6 sm:p-8 rounded-[2rem] shadow-xl shadow-indigo-200/50 text-white flex flex-col justify-between relative overflow-hidden">
          <div className="absolute -top-10 -right-10 w-40 h-40 bg-white opacity-10 rounded-full blur-2xl"></div>
          <div className="relative z-10">
            <div className="flex justify-between items-start mb-8">
              <div>
                <p className="text-indigo-100 font-bold mb-1 text-sm tracking-wide uppercase">Total Rendered</p>
                <h3 className="text-5xl font-black tracking-tighter">{totalHours.toFixed(1)}<span className="text-2xl opacity-70">h</span></h3>
              </div>
              <div className="bg-white/10 backdrop-blur-md p-2 rounded-2xl border border-white/20 flex flex-col items-center">
                <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-14 bg-transparent outline-none font-black text-center text-white text-lg" />
                <span className="text-[10px] font-bold text-indigo-100 uppercase tracking-wider">Goal</span>
              </div>
            </div>
            
            <div className="space-y-3 mb-6">
              <div className="flex justify-between text-sm font-bold text-indigo-100">
                <span>{progressPercentage >= 100 ? 'Goal Reached! 🎉' : `${remainingHours.toFixed(1)}h remaining`}</span>
                <span>{progressPercentage.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-black/20 rounded-full h-3 overflow-hidden border border-white/10 p-0.5">
                <div className="bg-gradient-to-r from-emerald-300 to-emerald-400 h-full rounded-full transition-all duration-1000 relative shadow-[0_0_10px_rgba(52,211,153,0.5)]" style={{ width: `${progressPercentage}%` }}></div>
              </div>
            </div>

            <div className="pt-4 border-t border-white/10 grid grid-cols-2 gap-4">
              <div className="bg-white/10 rounded-2xl p-3 backdrop-blur-sm">
                <p className="text-indigo-200 text-xs font-bold mb-1 uppercase tracking-wider">Days Present</p>
                <p className="font-black text-xl flex items-center gap-2"><CalendarDays className="w-5 h-5 text-indigo-300"/> {dtr.length}</p>
              </div>
              <div className="bg-white/10 rounded-2xl p-3 backdrop-blur-sm">
                <p className="text-indigo-200 text-xs font-bold mb-1 uppercase tracking-wider">Avg Hrs/Day</p>
                <p className="font-black text-xl flex items-center gap-2"><Clock className="w-5 h-5 text-indigo-300"/> {dtr.length > 0 ? (totalHours / dtr.length).toFixed(1) : 0}</p>
              </div>
            </div>
          </div>
        </div>
      </div>

      <DtrForm userId={session.user.id} onEntryAdded={fetchDTR} />

      {/* Modern Timesheet Section */}
      <div className="mt-8">
        {/* Toolbar */}
        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
           <div className="flex items-center gap-3">
              <h3 className="text-2xl font-black text-slate-800 tracking-tight">Timesheet History</h3>
              <span className="bg-slate-200 text-slate-700 text-xs font-bold px-3 py-1 rounded-full">{dtr.length} records</span>
           </div>
           
           <div className="relative w-full sm:w-72">
              <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-400" /></div>
              <input type="text" placeholder="Search remarks or date..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-3 bg-white border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-4 focus:ring-indigo-500/10 shadow-sm" />
           </div>
        </div>

        {/* Mobile View: Cards (Visible only on small screens) */}
        <div className="block md:hidden space-y-4">
          {filteredDtr.length === 0 ? (
            <div className="bg-white rounded-3xl p-8 text-center border border-slate-100">
               <SearchX className="w-12 h-12 text-slate-300 mx-auto mb-3" />
               <p className="text-slate-500 font-bold">No records found.</p>
            </div>
          ) : (
            filteredDtr.map((entry) => (
              <div key={entry.id} className="bg-white rounded-2xl p-5 border border-slate-100 shadow-sm relative overflow-hidden">
                 <div className="absolute top-0 left-0 w-1.5 h-full bg-indigo-500"></div>
                 <div className="flex justify-between items-start mb-3 pl-2">
                    <div>
                       <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">{new Date(entry.date).toLocaleDateString('en-US', { weekday:'short', month: 'short', day: 'numeric', year: 'numeric' })}</p>
                       <p className="font-black text-lg text-slate-800">{entry.time_in.substring(0, 5)} - {entry.time_out.substring(0, 5)}</p>
                    </div>
                    <span className="bg-emerald-100 text-emerald-700 font-black px-3 py-1 rounded-xl text-sm border border-emerald-200">{entry.hours_rendered}h</span>
                 </div>
                 <div className="bg-slate-50 rounded-xl p-3 mb-4 ml-2 border border-slate-100">
                    <p className="text-sm font-medium text-slate-600 line-clamp-2">{entry.remarks || <span className="italic text-slate-400">No remarks</span>}</p>
                 </div>
                 <div className="flex justify-end gap-2 border-t border-slate-100 pt-3 mt-1 pl-2">
                    <button onClick={() => handleEditClick(entry)} className="flex items-center gap-1.5 text-xs font-bold text-indigo-600 bg-indigo-50 px-4 py-2 rounded-xl active:scale-95"><Edit className="w-3.5 h-3.5"/> Edit</button>
                    <button onClick={() => deleteEntry(entry.id)} className="flex items-center gap-1.5 text-xs font-bold text-rose-600 bg-rose-50 px-4 py-2 rounded-xl active:scale-95"><Trash2 className="w-3.5 h-3.5"/> Delete</button>
                 </div>
              </div>
            ))
          )}
        </div>

        {/* Desktop View: Table (Visible only on medium screens and up) */}
        <div className="hidden md:block bg-white rounded-[2rem] shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-slate-100 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse whitespace-nowrap">
              <thead>
                <tr className="bg-slate-50/50 border-b border-slate-100">
                  <th className="p-5 text-xs font-black uppercase tracking-widest text-slate-400 w-40">Date</th>
                  <th className="p-5 text-xs font-black uppercase tracking-widest text-slate-400">Time Log</th>
                  <th className="p-5 text-xs font-black uppercase tracking-widest text-slate-400">Hours</th>
                  <th className="p-5 text-xs font-black uppercase tracking-widest text-slate-400 w-1/3">Remarks</th>
                  <th className="p-5 text-xs font-black uppercase tracking-widest text-slate-400 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {filteredDtr.length === 0 ? (
                  <tr><td colSpan="5" className="p-16 text-center text-slate-400 font-bold"><SearchX className="w-10 h-10 mx-auto mb-3 opacity-50"/> No timesheet records found.</td></tr>
                ) : (
                  filteredDtr.map((entry) => (
                    <tr key={entry.id} className="hover:bg-slate-50/50 transition-colors group">
                      <td className="p-5 font-bold text-slate-700">{new Date(entry.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}</td>
                      <td className="p-5">
                         <div className="flex items-center gap-2 font-black text-slate-800 bg-white border border-slate-200 px-3 py-1.5 rounded-xl w-max shadow-sm">
                            <span className="text-slate-500">{entry.time_in.substring(0, 5)}</span>
                            <span className="text-slate-300">→</span>
                            <span className="text-slate-500">{entry.time_out.substring(0, 5)}</span>
                         </div>
                      </td>
                      <td className="p-5"><span className="inline-flex items-center px-3 py-1 rounded-xl text-sm font-black bg-emerald-50 text-emerald-700 border border-emerald-200">{entry.hours_rendered}h</span></td>
                      <td className="p-5 text-slate-500 font-medium max-w-xs truncate">{entry.remarks || '-'}</td>
                      <td className="p-5 text-right">
                        <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                          <button onClick={() => handleEditClick(entry)} className="text-indigo-500 hover:text-indigo-700 hover:bg-indigo-50 p-2.5 rounded-xl transition-all border border-transparent hover:border-indigo-100"><Edit className="w-4 h-4"/></button>
                          <button onClick={() => deleteEntry(entry.id)} className="text-rose-500 hover:text-rose-700 hover:bg-rose-50 p-2.5 rounded-xl transition-all border border-transparent hover:border-rose-100"><Trash2 className="w-4 h-4"/></button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}