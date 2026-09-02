import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import DtrForm from './DtrForm';
import { LogOut, Clock, CalendarDays, Building2, GraduationCap, User, Trash2, Edit2, X, FileSpreadsheet, Search, SearchX, Briefcase, CheckCircle2 } from 'lucide-react';
import { calculateHoursRendered, formatDateString } from '../utils';

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
    <div className="w-full flex flex-col min-h-screen">
      
      {/* Edit Modal - True Responsive Design (No Overflowing) */}
      {editingEntry && (
        <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-slate-900/40 backdrop-blur-sm p-0 sm:p-4 transition-opacity">
          <div className="w-full sm:max-w-md bg-white rounded-t-3xl sm:rounded-2xl shadow-2xl flex flex-col max-h-[90vh] sm:max-h-auto">
            {/* Modal Header */}
            <div className="flex justify-between items-center p-5 border-b border-slate-100 flex-shrink-0">
              <h3 className="text-xl font-black text-slate-800">Edit Record</h3>
              <button onClick={() => setEditingEntry(null)} className="p-2 bg-slate-100 hover:bg-slate-200 rounded-full transition-colors"><X className="w-4 h-4 text-slate-600" /></button>
            </div>
            
            {/* Modal Scrollable Body */}
            <div className="p-5 overflow-y-auto space-y-4">
              <div>
                <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div className="w-full">
                  <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1 truncate">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-2 sm:px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
                </div>
                <div className="w-full">
                  <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1 truncate">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-2 sm:px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
                </div>
              </div>
              <div>
                <label className="block text-[11px] font-bold text-slate-500 uppercase tracking-wide mb-1.5 ml-1">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:bg-white focus:ring-2 focus:ring-indigo-500/20 font-bold outline-none text-slate-700" />
              </div>
              <label className="flex items-center space-x-3 cursor-pointer mt-2 bg-slate-50 p-3 rounded-xl border border-slate-200 select-none">
                <div className={`w-4 h-4 rounded flex items-center justify-center border transition-colors ${deductLunchEdit ? 'bg-indigo-600 border-indigo-600' : 'bg-white border-slate-300'}`}>
                   {deductLunchEdit && <CheckCircle2 className="w-3 h-3 text-white" />}
                </div>
                <input type="checkbox" checked={deductLunchEdit} onChange={(e) => setDeductLunchEdit(e.target.checked)} className="hidden" />
                <span className="text-sm font-bold text-slate-700">Auto-deduct 1h lunch break</span>
              </label>
            </div>

            {/* Modal Footer */}
            <div className="p-5 border-t border-slate-100 flex gap-3 flex-shrink-0 pb-safe">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold py-3 rounded-xl transition-colors active:scale-95">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-[2] bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 rounded-xl shadow-soft transition-all active:scale-95">
                {isSavingEdit ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Top Navigation */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div className="max-w-[1400px] mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3 min-w-0">
             <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center flex-shrink-0">
                <Briefcase className="w-4 h-4 text-white" />
             </div>
             <div className="min-w-0">
                <h1 className="text-sm sm:text-base font-black text-slate-900 truncate">
                  {profile.full_name || session.user.email.split('@')[0]}
                </h1>
                <p className="text-[10px] sm:text-xs font-bold text-slate-500 uppercase tracking-widest truncate">{profile.company_name || 'No Company'}</p>
             </div>
          </div>
          <div className="flex items-center gap-2 flex-shrink-0 ml-4">
            <button onClick={exportStyledExcel} className="flex items-center justify-center gap-1.5 bg-slate-900 hover:bg-indigo-600 text-white shadow-soft font-semibold px-3 py-2 sm:px-4 sm:py-2 rounded-lg transition-all active:scale-95 text-xs sm:text-sm">
              <FileSpreadsheet className="w-4 h-4" /> <span className="hidden sm:inline">Export</span>
            </button>
            <button onClick={() => supabase.auth.signOut()} className="flex items-center justify-center bg-slate-50 text-slate-500 hover:text-rose-600 hover:bg-rose-50 font-bold p-2 sm:px-3 sm:py-2 rounded-lg border border-slate-200 transition-all active:scale-95">
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </div>
      </header>
      
      {/* Main Content Area - Split Dashboard Layout for Desktop */}
      <main className="flex-1 w-full max-w-[1400px] mx-auto p-4 sm:p-6 lg:p-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 lg:gap-8">
          
          {/* LEFT COLUMN: Profile, Stats, Add Form */}
          <div className="lg:col-span-5 xl:col-span-4 space-y-6 flex flex-col">
            
            {/* Stats Overview */}
            <div className="bg-gradient-to-br from-indigo-600 via-violet-600 to-purple-700 p-6 rounded-2xl shadow-float text-white relative overflow-hidden group">
              <div className="absolute top-0 right-0 p-6 opacity-10 pointer-events-none"><Clock className="w-24 h-24"/></div>
              <div className="relative z-10">
                <div className="flex justify-between items-start mb-6">
                  <div>
                    <p className="text-indigo-200 font-bold mb-1 text-[10px] tracking-widest uppercase">Hours Rendered</p>
                    <h3 className="text-4xl font-black tracking-tighter drop-shadow-sm">{totalHours.toFixed(1)}<span className="text-xl opacity-70">h</span></h3>
                  </div>
                  <div className="bg-white/10 backdrop-blur-md px-2 py-1.5 rounded-xl border border-white/10 flex flex-col items-center">
                    <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-10 bg-transparent outline-none font-black text-center text-white text-sm" />
                    <span className="text-[8px] font-black text-indigo-200 uppercase tracking-widest mt-0.5">Goal</span>
                  </div>
                </div>
                
                <div className="space-y-2 mb-5">
                  <div className="flex justify-between text-[11px] font-bold text-indigo-100">
                    <span>{progressPercentage >= 100 ? 'Goal Reached! 🎉' : `${remainingHours.toFixed(1)}h left`}</span>
                    <span>{progressPercentage.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-black/20 rounded-full h-2.5 overflow-hidden">
                    <div className="bg-emerald-400 h-full rounded-full transition-all duration-1000" style={{ width: `${progressPercentage}%` }}></div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div className="bg-white/10 rounded-xl p-3">
                    <p className="text-indigo-200 text-[9px] font-bold mb-0.5 uppercase tracking-widest">Days Present</p>
                    <p className="font-black text-lg flex items-center gap-1.5"><CalendarDays className="w-3.5 h-3.5 text-indigo-300"/> {dtr.length}</p>
                  </div>
                  <div className="bg-white/10 rounded-xl p-3">
                    <p className="text-indigo-200 text-[9px] font-bold mb-0.5 uppercase tracking-widest">Avg Hours</p>
                    <p className="font-black text-lg flex items-center gap-1.5"><Clock className="w-3.5 h-3.5 text-indigo-300"/> {dtr.length > 0 ? (totalHours / dtr.length).toFixed(1) : 0}</p>
                  </div>
                </div>
              </div>
            </div>

            {/* Compact Profile Settings */}
            <div className="bg-white p-5 rounded-2xl shadow-soft border border-slate-100">
              <h2 className="text-sm font-black text-slate-800 mb-4 flex items-center gap-1.5"><User className="w-4 h-4 text-indigo-500"/> Profile Info</h2>
              <div className="space-y-3">
                <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 focus-within:bg-white focus-within:ring-2 focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all flex items-center gap-2">
                  <User className="w-4 h-4 text-slate-400 flex-shrink-0 ml-1"/>
                  <input type="text" defaultValue={profile.full_name || ''} placeholder="Full Name" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-700 text-sm min-w-0" />
                </div>
                <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 focus-within:bg-white focus-within:ring-2 focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all flex items-center gap-2">
                  <Briefcase className="w-4 h-4 text-slate-400 flex-shrink-0 ml-1"/>
                  <input type="text" defaultValue={profile.company_name} placeholder="Company Name" onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-700 text-sm min-w-0" />
                </div>
                <div className="bg-slate-50 p-2.5 rounded-xl border border-slate-200 focus-within:bg-white focus-within:ring-2 focus-within:ring-indigo-500/20 focus-within:border-indigo-300 transition-all flex items-center gap-2">
                  <GraduationCap className="w-4 h-4 text-slate-400 flex-shrink-0 ml-1"/>
                  <input type="text" defaultValue={profile.school} placeholder="School / University" onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-700 text-sm min-w-0" />
                </div>
              </div>
            </div>

            {/* DTR Form Component */}
            <DtrForm userId={session.user.id} onEntryAdded={fetchDTR} />
            
          </div>

          {/* RIGHT COLUMN: Timesheet History List / Table */}
          <div className="lg:col-span-7 xl:col-span-8 flex flex-col min-w-0">
            
            {/* Toolbar */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4 gap-3">
               <h3 className="text-xl font-black text-slate-800 flex items-center gap-2">
                 History <span className="bg-slate-200 text-slate-600 text-[10px] px-2 py-0.5 rounded-full">{dtr.length}</span>
               </h3>
               <div className="relative w-full sm:w-64">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-400" /></div>
                  <input type="text" placeholder="Search records..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} 
                    className="w-full pl-9 pr-3 py-2 bg-white border border-slate-200 rounded-lg text-sm font-bold text-slate-700 outline-none transition-all focus:ring-2 focus:ring-indigo-500/20 shadow-sm min-w-0" />
               </div>
            </div>

            {/* Timesheet Display */}
            {filteredDtr.length === 0 ? (
              <div className="bg-white rounded-2xl p-10 text-center border border-slate-100 shadow-soft flex-1 flex flex-col items-center justify-center">
                 <SearchX className="w-10 h-10 text-slate-300 mb-3" />
                 <p className="text-slate-500 font-bold text-sm">No timesheet records found.</p>
              </div>
            ) : (
              <div className="space-y-3">
                {/* Mobile Specific Stacked Cards (Visible <md) */}
                <div className="md:hidden space-y-3">
                  {filteredDtr.map((entry) => (
                    <div key={entry.id} className="bg-white rounded-2xl p-4 border border-slate-100 shadow-sm flex flex-col gap-3">
                       <div className="flex justify-between items-start min-w-0 gap-2">
                          <div className="min-w-0">
                             <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-0.5">{formatDateString(entry.date)}</p>
                             <p className="font-black text-lg text-slate-800 flex items-center gap-1.5 flex-wrap">
                               <span>{entry.time_in.substring(0, 5)}</span>
                               <span className="text-slate-300 text-xs">→</span>
                               <span>{entry.time_out.substring(0, 5)}</span>
                             </p>
                          </div>
                          <span className="bg-emerald-50 text-emerald-600 font-black px-2.5 py-1 rounded-lg text-xs border border-emerald-100 flex-shrink-0">{entry.hours_rendered}h</span>
                       </div>
                       
                       <div className="bg-slate-50 rounded-xl p-3 border border-slate-100">
                          <p className="text-xs font-medium text-slate-600 break-words">{entry.remarks || <span className="italic text-slate-400">No remarks added</span>}</p>
                       </div>
                       
                       <div className="flex justify-end gap-2">
                          <button onClick={() => handleEditClick(entry)} className="flex-1 flex justify-center items-center gap-1.5 text-xs font-bold text-slate-600 bg-white border border-slate-200 px-3 py-2 rounded-lg active:scale-95"><Edit2 className="w-3.5 h-3.5"/> Edit</button>
                          <button onClick={() => deleteEntry(entry.id)} className="flex items-center justify-center gap-1.5 text-xs font-bold text-rose-600 bg-white border border-rose-100 px-3 py-2 rounded-lg active:scale-95"><Trash2 className="w-3.5 h-3.5"/></button>
                       </div>
                    </div>
                  ))}
                </div>

                {/* Desktop Specific Table (Visible >=md) */}
                <div className="hidden md:block bg-white rounded-2xl shadow-soft border border-slate-100 overflow-hidden">
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-50/80 border-b border-slate-100">
                          <th className="p-4 px-5 text-[10px] font-bold uppercase tracking-widest text-slate-400 whitespace-nowrap">Date</th>
                          <th className="p-4 text-[10px] font-bold uppercase tracking-widest text-slate-400 whitespace-nowrap">Time Log</th>
                          <th className="p-4 text-[10px] font-bold uppercase tracking-widest text-slate-400 whitespace-nowrap">Hrs</th>
                          <th className="p-4 text-[10px] font-bold uppercase tracking-widest text-slate-400 w-full">Remarks</th>
                          <th className="p-4 px-5 text-[10px] font-bold uppercase tracking-widest text-slate-400 text-right whitespace-nowrap">Act</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-50">
                        {filteredDtr.map((entry) => (
                          <tr key={entry.id} className="hover:bg-slate-50/50 transition-colors group">
                            <td className="p-4 px-5 font-bold text-slate-700 text-sm whitespace-nowrap">{formatDateString(entry.date)}</td>
                            <td className="p-4 whitespace-nowrap">
                               <div className="flex items-center gap-2 font-black text-slate-800 bg-white border border-slate-200 px-2.5 py-1.5 rounded-lg w-max shadow-sm text-sm">
                                  <span>{entry.time_in.substring(0, 5)}</span>
                                  <span className="text-slate-300">→</span>
                                  <span>{entry.time_out.substring(0, 5)}</span>
                               </div>
                            </td>
                            <td className="p-4 whitespace-nowrap"><span className="inline-flex items-center px-2.5 py-1 rounded-lg text-xs font-black bg-emerald-50 text-emerald-600 border border-emerald-100">{entry.hours_rendered}h</span></td>
                            <td className="p-4 text-slate-600 font-medium text-sm break-words min-w-[200px]">{entry.remarks || <span className="text-slate-300 italic">No remarks</span>}</td>
                            <td className="p-4 px-5 text-right whitespace-nowrap">
                              <div className="flex justify-end gap-1.5 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button onClick={() => handleEditClick(entry)} className="text-slate-400 hover:text-indigo-600 bg-white hover:bg-indigo-50 p-2 rounded-lg transition-all border border-slate-200 hover:border-indigo-100"><Edit2 className="w-3.5 h-3.5"/></button>
                                <button onClick={() => deleteEntry(entry.id)} className="text-slate-400 hover:text-rose-600 bg-white hover:bg-rose-50 p-2 rounded-lg transition-all border border-slate-200 hover:border-rose-100"><Trash2 className="w-3.5 h-3.5"/></button>
                              </div>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}