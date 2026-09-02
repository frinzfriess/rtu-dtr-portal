import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import DtrForm from './DtrForm';
import { LogOut, Clock, CalendarDays, Building2, GraduationCap, User, Trash2, Edit2, X, FileSpreadsheet, Search, SearchX, Briefcase } from 'lucide-react';
import { calculateHoursRendered, getGreeting, formatDateString } from '../utils';

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
            company_name: 'Pending Assignment', required_hours: 300, full_name: '' 
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
    if(!window.confirm("Delete this record permanently?")) return;
    try {
      await supabase.from('dtr_entries').delete().eq('id', id);
      fetchDTR();
    } catch (error) { alert("Could not delete record."); }
  }

  const handleEditClick = (entry) => {
    setEditingEntry(entry);
    setEditForm({
      date: entry.date, time_in: entry.time_in.substring(0, 5), time_out: entry.time_out.substring(0, 5), remarks: entry.remarks || ''
    });
    setDeductLunchEdit(entry.hours_rendered < 5 ? false : true); 
  };

  const saveEdit = async () => {
    setIsSavingEdit(true);
    const newHours = calculateHoursRendered(editForm.time_in, editForm.time_out, deductLunchEdit);
    if (newHours <= 0) { alert("Invalid time range."); setIsSavingEdit(false); return; }
    try {
      const { error } = await supabase.from('dtr_entries').update({
        date: editForm.date, time_in: `${editForm.time_in}:00`, time_out: `${editForm.time_out}:00`, hours_rendered: newHours, remarks: editForm.remarks
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
      const headerRow = worksheet.addRow(['Date', 'Time In', 'Time Out', 'Hours', 'Remarks']);
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
    <div className="w-full relative pb-24 sm:pb-12 bg-[#F8FAFC]">
      
      {/* Dynamic Edit Modal (Mobile Bottom Sheet / Desktop Modal) */}
      {editingEntry && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-end sm:items-center justify-center z-[100] transition-all p-0 sm:p-4">
          <div className="bg-white w-full sm:max-w-md rounded-t-[2rem] sm:rounded-[2rem] p-6 sm:p-8 shadow-2xl relative border border-slate-100 animate-[slideUp_0.3s_ease-out] sm:animate-none">
            <div className="w-12 h-1.5 bg-slate-200 rounded-full mx-auto mb-6 sm:hidden"></div>
            
            <button onClick={() => setEditingEntry(null)} className="absolute top-6 right-6 text-slate-400 hover:text-slate-700 bg-slate-50 hover:bg-slate-100 p-2 rounded-full transition-all"><X className="w-5 h-5" /></button>
            <h3 className="text-2xl font-black text-slate-800 mb-6 flex items-center gap-2">Edit Record</h3>
            
            <div className="space-y-4">
              <div>
                <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
              </div>
              <div className="flex gap-4">
                <div className="flex-1">
                  <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
                </div>
                <div className="flex-1">
                  <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
                </div>
              </div>
              <div>
                <label className="block text-[11px] font-bold text-slate-400 uppercase tracking-widest mb-1.5 ml-2">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-4 py-3.5 bg-slate-50 border border-slate-200 rounded-2xl focus:bg-white focus:ring-[3px] focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
              </div>
              <label className="flex items-center space-x-3 cursor-pointer mt-2 bg-slate-50 p-4 rounded-2xl border border-slate-200 select-none">
                <div className={`w-5 h-5 rounded flex items-center justify-center border transition-colors ${deductLunchEdit ? 'bg-indigo-600 border-indigo-600' : 'bg-white border-slate-300'}`}>
                   {deductLunchEdit && <CheckCircle2 className="w-3.5 h-3.5 text-white" />}
                </div>
                <input type="checkbox" checked={deductLunchEdit} onChange={(e) => setDeductLunchEdit(e.target.checked)} className="hidden" />
                <span className="text-sm font-bold text-slate-700">Auto-deduct 1h lunch break</span>
              </label>
            </div>
            <div className="flex gap-3 mt-8 pb-4 sm:pb-0">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-4 rounded-2xl transition-colors active:scale-95">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-[2] bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-4 rounded-2xl shadow-float transition-all active:scale-95">
                {isSavingEdit ? 'Saving...' : 'Save Changes'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Top Header / Sticky Nav */}
      <div className="sticky top-0 z-40 bg-[#F8FAFC]/80 backdrop-blur-xl border-b border-slate-200/50 px-4 sm:px-6 lg:px-8 py-4 mb-6 sm:mb-8 transition-all">
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <div className="flex flex-col">
            <p className="text-indigo-600 font-bold text-[10px] sm:text-xs tracking-widest uppercase mb-0.5">{getGreeting()}</p>
            <h1 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight leading-none truncate max-w-[200px] sm:max-w-md">
              {profile.full_name || session.user.email.split('@')[0]}
            </h1>
          </div>
          <div className="flex gap-2 sm:gap-3">
            <button onClick={exportStyledExcel} className="flex items-center justify-center gap-2 bg-slate-900 hover:bg-indigo-600 text-white shadow-[0_4px_15px_rgb(0,0,0,0.1)] hover:shadow-float font-semibold px-4 py-2.5 sm:px-5 sm:py-2.5 rounded-xl sm:rounded-2xl transition-all active:scale-95 text-xs sm:text-sm">
              <FileSpreadsheet className="w-4 h-4" /> <span className="hidden sm:block">Export DTR</span>
            </button>
            <button onClick={() => supabase.auth.signOut()} className="flex items-center justify-center bg-white text-slate-500 hover:text-rose-600 hover:bg-rose-50 font-bold p-2.5 sm:px-4 sm:py-2.5 rounded-xl sm:rounded-2xl border border-slate-200 transition-all active:scale-95 shadow-sm">
              <LogOut className="w-4 h-4 sm:w-5 sm:h-5" />
            </button>
          </div>
        </div>
      </div>
      
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        
        {/* Profile & Stats Layout */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
          
          {/* Left Side: Profile Cards */}
          <div className="lg:col-span-2 bg-white p-5 sm:p-8 rounded-[2rem] shadow-soft border border-slate-100/50">
            <h2 className="text-lg font-black text-slate-800 mb-6 flex items-center gap-2"><User className="w-5 h-5 text-indigo-500"/> Student Profile</h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <div className="bg-slate-50 p-4 rounded-[1.5rem] border border-slate-200 focus-within:bg-white focus-within:ring-[3px] focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all">
                <label className="text-[10px] uppercase tracking-widest text-slate-400 font-bold flex items-center gap-1.5 mb-1.5"><User className="w-3 h-3"/> Full Name</label>
                <input type="text" defaultValue={profile.full_name || ''} placeholder="Add your name..." onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-800 text-base sm:text-lg" />
              </div>
              <div className="bg-slate-50 p-4 rounded-[1.5rem] border border-slate-200 focus-within:bg-white focus-within:ring-[3px] focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all">
                <label className="text-[10px] uppercase tracking-widest text-slate-400 font-bold flex items-center gap-1.5 mb-1.5"><Briefcase className="w-3 h-3"/> Company</label>
                <input type="text" defaultValue={profile.company_name} placeholder="Add company name..." onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-800 text-base sm:text-lg" />
              </div>
              <div className="bg-slate-50 p-4 rounded-[1.5rem] border border-slate-200 focus-within:bg-white focus-within:ring-[3px] focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all sm:col-span-2">
                <label className="text-[10px] uppercase tracking-widest text-slate-400 font-bold flex items-center gap-1.5 mb-1.5"><GraduationCap className="w-3 h-3"/> School / University</label>
                <input type="text" defaultValue={profile.school} onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-800 text-base sm:text-lg" />
              </div>
            </div>
          </div>

          {/* Right Side: Stunning Stats Widget */}
          <div className="bg-slate-900 bg-[url('https://www.transparenttextures.com/patterns/cubes.png')] p-6 sm:p-8 rounded-[2rem] shadow-soft text-white flex flex-col justify-between relative overflow-hidden group">
            {/* Animated Gradient Background */}
            <div className="absolute inset-0 bg-gradient-to-br from-indigo-600/90 via-violet-600/90 to-purple-800/90 opacity-100 z-0"></div>
            <div className="absolute -top-20 -right-20 w-64 h-64 bg-white/20 rounded-full blur-[80px] z-0 group-hover:bg-white/30 transition-all duration-700"></div>
            
            <div className="relative z-10 h-full flex flex-col justify-between">
              <div className="flex justify-between items-start mb-6">
                <div>
                  <p className="text-indigo-200/80 font-bold mb-1 text-[11px] tracking-widest uppercase">Hours Rendered</p>
                  <h3 className="text-5xl sm:text-6xl font-black tracking-tighter text-white drop-shadow-sm">{totalHours.toFixed(1)}<span className="text-2xl sm:text-3xl opacity-60 font-bold">h</span></h3>
                </div>
                <div className="bg-white/10 backdrop-blur-md p-2.5 rounded-2xl border border-white/10 flex flex-col items-center shadow-inner">
                  <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-12 bg-transparent outline-none font-black text-center text-white text-lg" />
                  <span className="text-[9px] font-black text-indigo-200/80 uppercase tracking-widest mt-0.5">Goal</span>
                </div>
              </div>
              
              <div className="space-y-3 mb-6">
                <div className="flex justify-between text-xs font-bold text-indigo-100">
                  <span>{progressPercentage >= 100 ? 'Goal Reached! 🎉' : `${remainingHours.toFixed(1)}h remaining`}</span>
                  <span>{progressPercentage.toFixed(1)}%</span>
                </div>
                <div className="w-full bg-black/30 rounded-full h-3 overflow-hidden p-0.5 shadow-inner">
                  <div className="bg-gradient-to-r from-emerald-400 to-cyan-400 h-full rounded-full transition-all duration-1000 relative shadow-[0_0_15px_rgba(52,211,153,0.4)]" style={{ width: `${progressPercentage}%` }}></div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 mt-auto">
                <div className="bg-white/10 rounded-2xl p-3.5 backdrop-blur-md border border-white/5">
                  <p className="text-indigo-200/80 text-[10px] font-bold mb-1 uppercase tracking-widest">Days Present</p>
                  <p className="font-black text-xl flex items-center gap-2"><CalendarDays className="w-4 h-4 text-indigo-300"/> {dtr.length}</p>
                </div>
                <div className="bg-white/10 rounded-2xl p-3.5 backdrop-blur-md border border-white/5">
                  <p className="text-indigo-200/80 text-[10px] font-bold mb-1 uppercase tracking-widest">Avg / Day</p>
                  <p className="font-black text-xl flex items-center gap-2"><Clock className="w-4 h-4 text-indigo-300"/> {dtr.length > 0 ? (totalHours / dtr.length).toFixed(1) : 0}<span className="text-sm opacity-70">h</span></p>
                </div>
              </div>
            </div>
          </div>
        </div>

        <DtrForm userId={session.user.id} onEntryAdded={fetchDTR} />

        {/* Timesheet Section */}
        <div className="mt-8 mb-4">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 gap-4">
             <div className="flex items-center gap-3">
                <h3 className="text-2xl font-black text-slate-800 tracking-tight">Timesheet History</h3>
                <span className="bg-slate-200 text-slate-600 text-xs font-bold px-3 py-1 rounded-full">{dtr.length} records</span>
             </div>
             
             <div className="relative w-full sm:w-72">
                <div className="absolute inset-y-0 left-0 pl-4 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-400" /></div>
                <input type="text" placeholder="Search date or remarks..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} 
                  className="w-full pl-11 pr-4 py-3 bg-white border border-slate-200 rounded-2xl text-sm font-bold text-slate-700 outline-none transition-all focus:ring-[3px] focus:ring-indigo-500/20 shadow-sm" />
             </div>
          </div>

          {/* Mobile View: High-end Cards */}
          <div className="block md:hidden space-y-4">
            {filteredDtr.length === 0 ? (
              <div className="bg-white rounded-[2rem] p-8 text-center border border-slate-100 shadow-sm">
                 <SearchX className="w-12 h-12 text-slate-300 mx-auto mb-3" />
                 <p className="text-slate-500 font-bold">No records found.</p>
              </div>
            ) : (
              filteredDtr.map((entry) => (
                <div key={entry.id} className="bg-white rounded-[2rem] p-5 border border-slate-100 shadow-sm relative overflow-hidden flex flex-col gap-3">
                   <div className="flex justify-between items-start">
                      <div>
                         <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">{formatDateString(entry.date)}</p>
                         <p className="font-black text-xl text-slate-800 flex items-center gap-2">
                           {entry.time_in.substring(0, 5)} <span className="text-slate-300 text-sm">→</span> {entry.time_out.substring(0, 5)}
                         </p>
                      </div>
                      <span className="bg-emerald-50 text-emerald-600 font-black px-3 py-1.5 rounded-xl text-sm border border-emerald-100 shadow-sm">{entry.hours_rendered}h</span>
                   </div>
                   
                   <div className="bg-[#F8FAFC] rounded-2xl p-3.5 border border-slate-100/50">
                      <p className="text-sm font-medium text-slate-600 line-clamp-2">{entry.remarks || <span className="italic text-slate-400">No remarks added</span>}</p>
                   </div>
                   
                   <div className="flex justify-end gap-2 pt-1">
                      <button onClick={() => handleEditClick(entry)} className="flex-1 flex justify-center items-center gap-1.5 text-xs font-bold text-slate-600 bg-slate-50 border border-slate-200 px-4 py-3 rounded-xl active:scale-95 transition-all"><Edit2 className="w-4 h-4"/> Edit</button>
                      <button onClick={() => deleteEntry(entry.id)} className="flex items-center justify-center gap-1.5 text-xs font-bold text-rose-600 bg-rose-50 border border-rose-100 px-4 py-3 rounded-xl active:scale-95 transition-all"><Trash2 className="w-4 h-4"/></button>
                   </div>
                </div>
              ))
            )}
          </div>

          {/* Desktop View: Premium Table */}
          <div className="hidden md:block bg-white rounded-[2rem] shadow-soft border border-slate-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse whitespace-nowrap">
                <thead>
                  <tr className="bg-slate-50/80 border-b border-slate-100">
                    <th className="p-5 px-6 text-[10px] font-bold uppercase tracking-widest text-slate-400 w-48">Date</th>
                    <th className="p-5 text-[10px] font-bold uppercase tracking-widest text-slate-400 w-64">Time Log</th>
                    <th className="p-5 text-[10px] font-bold uppercase tracking-widest text-slate-400 w-32">Hours</th>
                    <th className="p-5 text-[10px] font-bold uppercase tracking-widest text-slate-400">Remarks</th>
                    <th className="p-5 px-6 text-[10px] font-bold uppercase tracking-widest text-slate-400 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  {filteredDtr.length === 0 ? (
                    <tr><td colSpan="5" className="p-16 text-center text-slate-400 font-bold"><SearchX className="w-12 h-12 mx-auto mb-3 opacity-30"/> No records found in your timesheet.</td></tr>
                  ) : (
                    filteredDtr.map((entry) => (
                      <tr key={entry.id} className="hover:bg-slate-50/50 transition-colors group">
                        <td className="p-5 px-6 font-bold text-slate-700">{formatDateString(entry.date)}</td>
                        <td className="p-5">
                           <div className="flex items-center gap-2.5 font-black text-slate-800 bg-white border border-slate-200 px-3.5 py-2 rounded-xl w-max shadow-sm">
                              <span className="text-slate-600">{entry.time_in.substring(0, 5)}</span>
                              <span className="text-slate-300">→</span>
                              <span className="text-slate-600">{entry.time_out.substring(0, 5)}</span>
                           </div>
                        </td>
                        <td className="p-5"><span className="inline-flex items-center px-3 py-1.5 rounded-xl text-sm font-black bg-emerald-50 text-emerald-600 border border-emerald-100 shadow-sm">{entry.hours_rendered}h</span></td>
                        <td className="p-5 text-slate-600 font-medium max-w-sm truncate">{entry.remarks || <span className="text-slate-300 italic">No remarks</span>}</td>
                        <td className="p-5 px-6 text-right">
                          <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                            <button onClick={() => handleEditClick(entry)} className="text-slate-400 hover:text-indigo-600 bg-white hover:bg-indigo-50 p-2.5 rounded-xl transition-all border border-slate-200 hover:border-indigo-100 shadow-sm"><Edit2 className="w-4 h-4"/></button>
                            <button onClick={() => deleteEntry(entry.id)} className="text-slate-400 hover:text-rose-600 bg-white hover:bg-rose-50 p-2.5 rounded-xl transition-all border border-slate-200 hover:border-rose-100 shadow-sm"><Trash2 className="w-4 h-4"/></button>
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
    </div>
  );
}