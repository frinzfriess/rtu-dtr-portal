import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import DtrForm from './DtrForm';
import { LogOut, Clock, CalendarDays, Building2, GraduationCap, User, Trash2, Edit2, X, FileSpreadsheet, Search, SearchX, ShieldCheck, Activity, Cpu, ShieldAlert, Wifi } from 'lucide-react';
import { calculateHoursRendered, formatDateString, getGreeting } from '../utils';

export default function Dashboard({ session, darkMode }) {
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
            company_name: 'Cyber Security Operations Center', required_hours: 300, full_name: '' 
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
    } catch (error) { alert('Failed to update clearance profile'); }
  }

  async function deleteEntry(id) {
    if(!window.confirm("Purge this telemetry record?")) return;
    try {
      await supabase.from('dtr_entries').delete().eq('id', id);
      fetchDTR();
    } catch (error) { alert("Could not purge record."); }
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
    if (newHours <= 0) { alert("Invalid time vector."); setIsSavingEdit(false); return; }
    try {
      const { error } = await supabase.from('dtr_entries').update({
        date: editForm.date, time_in: `${editForm.time_in}:00`, time_out: `${editForm.time_out}:00`, hours_rendered: newHours, remarks: editForm.remarks
      }).eq('id', editingEntry.id);
      if (error) throw error;
      setEditingEntry(null); fetchDTR();
    } catch (err) { alert("Failed to modify record."); } 
    finally { setIsSavingEdit(false); }
  };

  const exportStyledExcel = async () => {
    try {
      const ExcelJS = (await import('exceljs')).default || await import('exceljs');
      const { saveAs } = (await import('file-saver')).default || await import('file-saver');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('CyberSec DTR');
      worksheet.columns = [ { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 40 } ];

      worksheet.mergeCells('A1:E1');
      const title = worksheet.getCell('A1'); title.value = 'RIZAL TECHNOLOGICAL UNIVERSITY'; title.font = { bold: true, size: 14 }; title.alignment = { horizontal: 'center' };
      worksheet.mergeCells('A2:E2');
      const subTitle = worksheet.getCell('A2'); subTitle.value = 'CYBERSECURITY OJT - DAILY TIME RECORD'; subTitle.font = { bold: true, size: 12 }; subTitle.alignment = { horizontal: 'center' };
      worksheet.addRow([]);
      worksheet.addRow(['Operative:', profile.full_name || 'N/A', '', 'Station:', profile.company_name || 'N/A']);
      worksheet.addRow(['Track:', profile.course || 'N/A', '', 'Institution:', profile.school || 'N/A']);
      worksheet.addRow([]);
      const headerRow = worksheet.addRow(['Date', 'Time In', 'Time Out', 'Hours', 'Incident Remarks']);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } }; headerRow.alignment = { horizontal: 'center' };
      headerRow.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });

      const sortedDtr = [...dtr].sort((a,b) => new Date(a.date) - new Date(b.date));
      sortedDtr.forEach(entry => {
          const row = worksheet.addRow([ new Date(entry.date).toLocaleDateString(), entry.time_in.substring(0, 5), entry.time_out.substring(0, 5), entry.hours_rendered, entry.remarks || '' ]);
          row.eachCell(cell => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' }; });
      });
      worksheet.addRow([]);
      const totalRow = worksheet.addRow(['', '', 'TOTAL HOURS:', totalHours.toFixed(2), '']);
      totalRow.font = { bold: true, size: 12 }; totalRow.getCell(4).font = { bold: true, size: 12, color: { argb: 'FF06B6D4' } };
      
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `CyberDTR_${profile.full_name || 'Operative'}.xlsx`);
    } catch (err) { alert("Excel export failed."); }
  };

  if (loading || !profile) return null;
  const progressPercentage = Math.min((totalHours / profile.required_hours) * 100, 100);
  const remainingHours = Math.max(profile.required_hours - totalHours, 0);
  const filteredDtr = dtr.filter(e => e.remarks?.toLowerCase().includes(searchTerm.toLowerCase()) || e.date.includes(searchTerm));

  return (
    <div className="w-full flex flex-col min-h-screen transition-colors duration-300">
      
      {/* Edit Modal - Adaptive Theme */}
      {editingEntry && (
        <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-slate-950/70 backdrop-blur-sm p-0 sm:p-4 transition-opacity">
          <div className="w-full sm:max-w-md bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-t-3xl sm:rounded-2xl shadow-2xl flex flex-col max-h-[90vh]">
            <div className="flex justify-between items-center p-4 sm:p-5 border-b border-slate-100 dark:border-slate-800 flex-shrink-0">
              <h3 className="text-sm font-black text-slate-900 dark:text-white uppercase tracking-wider font-mono">Modify Telemetry Record</h3>
              <button onClick={() => setEditingEntry(null)} className="p-2 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-full transition-colors text-slate-500 dark:text-slate-400"><X className="w-4 h-4" /></button>
            </div>
            
            <div className="p-4 sm:p-5 overflow-y-auto space-y-3.5 font-mono text-xs">
              <div>
                <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1">Date Vector</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500" />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div className="w-full">
                  <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500" />
                </div>
                <div className="w-full">
                  <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-3 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500" />
                </div>
              </div>
              <div>
                <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 ml-1 font-sans">Incident Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-3.5 py-2.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-sans font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500" />
              </div>
              <label className="flex items-center space-x-3 cursor-pointer mt-1 bg-slate-50 dark:bg-slate-950 p-3 rounded-xl border border-slate-200 dark:border-slate-800 select-none">
                <div className={`w-4 h-4 rounded flex items-center justify-center border transition-colors ${deductLunchEdit ? 'bg-cyan-500 border-cyan-400' : 'bg-transparent border-slate-400'}`}>
                   {deductLunchEdit && <span className="text-[10px] text-slate-950 font-bold">✓</span>}
                </div>
                <input type="checkbox" checked={deductLunchEdit} onChange={(e) => setDeductLunchEdit(e.target.checked)} className="hidden" />
                <span className="font-bold text-slate-700 dark:text-slate-300">Auto-deduct 1h lunch break</span>
              </label>
            </div>

            <div className="p-4 sm:p-5 border-t border-slate-100 dark:border-slate-800 flex gap-3 flex-shrink-0">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 font-bold py-3 rounded-xl transition-colors font-mono text-xs">Abort</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-[2] bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black py-3 rounded-xl shadow-glow transition-all font-mono text-xs">
                {isSavingEdit ? 'Encrypting...' : 'Commit Changes'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Top Cyber Security Header Bar */}
      <header className="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 sticky top-0 z-30 transition-colors">
        <div className="max-w-[1400px] mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2.5 min-w-0">
             <div className="w-8 h-8 rounded-xl bg-cyan-500/10 border border-cyan-500/30 flex items-center justify-center flex-shrink-0 text-cyan-600 dark:text-cyan-400 shadow-glow">
                <ShieldCheck className="w-4 h-4" />
             </div>
             <div className="min-w-0">
                <div className="flex items-center gap-2">
                   <h1 className="text-xs sm:text-sm font-black text-slate-900 dark:text-white tracking-wide truncate">
                     {profile.full_name || session.user.email.split('@')[0]}
                   </h1>
                   <span className="hidden sm:inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-mono font-bold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">SOC-SECURE</span>
                </div>
                <p className="text-[9px] font-mono text-cyan-600 dark:text-cyan-400 truncate">{profile.company_name || 'Cyber Operations Center'}</p>
             </div>
          </div>
          <div className="flex items-center gap-2 flex-shrink-0 ml-2 font-mono">
            <button onClick={exportStyledExcel} className="flex items-center justify-center gap-1.5 bg-cyan-500 hover:bg-cyan-400 text-slate-950 shadow-glow font-black px-3 py-2 rounded-xl transition-all active:scale-95 text-xs">
              <FileSpreadsheet className="w-3.5 h-3.5" /> <span className="hidden sm:inline">Export DTR</span>
            </button>
            <button onClick={() => supabase.auth.signOut()} className="flex items-center justify-center bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:text-rose-600 dark:hover:text-rose-400 hover:bg-rose-500/10 font-bold p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 transition-all active:scale-95">
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </div>
      </header>
      
      {/* Main Dashboard Grid */}
      <main className="flex-1 w-full max-w-[1400px] mx-auto p-4 sm:p-6 lg:p-8">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 lg:gap-6">
          
          {/* LEFT COLUMN: Controls, Stats, Profile, Form */}
          <div className="lg:col-span-5 xl:col-span-4 space-y-5 flex flex-col">
            
            {/* Cyber Threat / Progress Widget */}
            <div className="bg-gradient-to-br from-slate-900 via-indigo-950 to-slate-900 border border-cyan-500/30 p-5 rounded-2xl shadow-xl text-white relative overflow-hidden">
              <div className="absolute top-0 right-0 p-5 opacity-10 pointer-events-none"><Activity className="w-20 h-20 text-cyan-400 animate-pulse"/></div>
              <div className="relative z-10">
                <div className="flex justify-between items-start mb-4">
                  <div>
                    <div className="flex items-center gap-1.5 text-cyan-400 text-[9px] font-mono font-bold uppercase tracking-widest mb-1">
                      <span className="w-2 h-2 rounded-full bg-cyan-400 animate-ping"></span>
                      Hours Rendered
                    </div>
                    <h3 className="text-3xl font-black tracking-tighter text-white font-mono">{totalHours.toFixed(1)}<span className="text-lg opacity-60 font-sans">h</span></h3>
                  </div>
                  <div className="bg-slate-950/80 border border-slate-800 px-2.5 py-1.5 rounded-xl flex flex-col items-center">
                    <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-10 bg-transparent outline-none font-black text-center text-cyan-400 text-xs font-mono" />
                    <span className="text-[7px] font-bold text-slate-400 uppercase tracking-widest mt-0.5 font-mono">Target</span>
                  </div>
                </div>
                
                <div className="space-y-1.5 mb-4">
                  <div className="flex justify-between text-[10px] font-bold font-mono text-slate-300">
                    <span>{progressPercentage >= 100 ? 'TARGET SECURED 🎉' : `${remainingHours.toFixed(1)}h remaining`}</span>
                    <span className="text-cyan-400">{progressPercentage.toFixed(1)}%</span>
                  </div>
                  <div className="w-full bg-slate-950 rounded-full h-2 overflow-hidden border border-slate-800 p-0.5">
                    <div className="bg-gradient-to-r from-cyan-400 to-emerald-400 h-full rounded-full transition-all duration-1000 shadow-glow" style={{ width: `${progressPercentage}%` }}></div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2.5 font-mono">
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-2.5">
                    <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-0.5">Shifts Logged</p>
                    <p className="font-black text-sm text-white flex items-center gap-1"><CalendarDays className="w-3 h-3 text-cyan-400"/> {dtr.length}</p>
                  </div>
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-2.5">
                    <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-0.5">Avg / Shift</p>
                    <p className="font-black text-sm text-white flex items-center gap-1"><Clock className="w-3 h-3 text-cyan-400"/> {dtr.length > 0 ? (totalHours / dtr.length).toFixed(1) : 0}h</p>
                  </div>
                </div>
              </div>
            </div>

            {/* Operative Clearance Profile */}
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-4 sm:p-5 rounded-2xl shadow-sm transition-colors">
              <h2 className="text-xs font-black text-slate-900 dark:text-white uppercase tracking-widest mb-3 flex items-center gap-2 font-mono"><User className="w-3.5 h-3.5 text-cyan-600 dark:text-cyan-400"/> Operative Profile</h2>
              <div className="space-y-2.5">
                <div className="bg-slate-50 dark:bg-slate-950 p-2.5 rounded-xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all flex items-center gap-2">
                  <User className="w-3.5 h-3.5 text-slate-400 flex-shrink-0 ml-0.5"/>
                  <input type="text" defaultValue={profile.full_name || ''} placeholder="Operative Full Name" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-xs min-w-0" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-2.5 rounded-xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all flex items-center gap-2">
                  <Building2 className="w-3.5 h-3.5 text-slate-400 flex-shrink-0 ml-0.5"/>
                  <input type="text" defaultValue={profile.company_name} placeholder="Assigned Cyber Station" onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-xs min-w-0" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-2.5 rounded-xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all flex items-center gap-2">
                  <GraduationCap className="w-3.5 h-3.5 text-slate-400 flex-shrink-0 ml-0.5"/>
                  <input type="text" defaultValue={profile.school} placeholder="University / Institution" onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-xs min-w-0" />
                </div>
              </div>
            </div>

            {/* Telemetry DTR Form */}
            <DtrForm userId={session.user.id} onEntryAdded={fetchDTR} />
            
          </div>

          {/* RIGHT COLUMN: Telemetry History */}
          <div className="lg:col-span-7 xl:col-span-8 flex flex-col min-w-0">
            
            {/* Toolbar */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-3.5 gap-3">
               <h3 className="text-sm font-black text-slate-900 dark:text-white uppercase tracking-wider flex items-center gap-2 font-mono">
                 Telemetry Logs <span className="bg-cyan-500/10 text-cyan-600 dark:text-cyan-400 border border-cyan-500/20 text-[10px] px-2 py-0.5 rounded-full">{dtr.length}</span>
               </h3>
               <div className="relative w-full sm:w-60">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none"><Search className="h-3.5 w-3.5 text-slate-400" /></div>
                  <input type="text" placeholder="Search logs..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} 
                    className="w-full pl-9 pr-3 py-2 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 shadow-sm min-w-0 font-mono" />
               </div>
            </div>

            {/* History Feed */}
            {filteredDtr.length === 0 ? (
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-10 text-center shadow-sm flex-1 flex flex-col items-center justify-center transition-colors">
                 <SearchX className="w-10 h-10 text-slate-400 dark:text-slate-600 mb-3" />
                 <p className="text-slate-500 dark:text-slate-400 font-mono text-xs">No telemetry logs found matching filter.</p>
              </div>
            ) : (
              <div className="space-y-3">
                {/* Mobile Responsive Cards */}
                <div className="space-y-3">
                  {filteredDtr.map((entry) => (
                    <div key={entry.id} className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-2xl p-4 shadow-sm flex flex-col gap-2.5 relative overflow-hidden transition-colors">
                       <div className="absolute top-0 left-0 w-1.5 h-full bg-cyan-500"></div>
                       <div className="flex justify-between items-start min-w-0 gap-2 pl-2">
                          <div className="min-w-0">
                             <p className="text-[10px] font-mono font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-0.5">{formatDateString(entry.date)}</p>
                             <p className="font-mono font-black text-sm text-slate-900 dark:text-white flex items-center gap-1.5 flex-wrap">
                               <span>{entry.time_in.substring(0, 5)}</span>
                               <span className="text-slate-400 text-xs">→</span>
                               <span>{entry.time_out.substring(0, 5)}</span>
                             </p>
                          </div>
                          <span className="bg-cyan-500/10 text-cyan-700 dark:text-cyan-300 border border-cyan-500/20 font-mono font-black px-2.5 py-1 rounded-lg text-xs flex-shrink-0">{entry.hours_rendered}h</span>
                       </div>
                       
                       <div className="bg-slate-50 dark:bg-slate-950 rounded-xl p-2.5 border border-slate-200 dark:border-slate-800 ml-2">
                          <p className="text-xs font-sans text-slate-600 dark:text-slate-300 break-words">{entry.remarks || <span className="italic text-slate-400">No incident remarks</span>}</p>
                       </div>
                       
                       <div className="flex justify-end gap-2 pl-2 pt-1">
                          <button onClick={() => handleEditClick(entry)} className="flex-1 flex justify-center items-center gap-1.5 text-xs font-bold font-mono text-slate-700 dark:text-slate-300 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 px-3 py-2 rounded-xl active:scale-95 transition-all">Edit</button>
                          <button onClick={() => deleteEntry(entry.id)} className="flex items-center justify-center gap-1.5 text-xs font-bold font-mono text-rose-600 dark:text-rose-400 bg-rose-500/10 hover:bg-rose-500/20 px-3 py-2 rounded-xl border border-rose-500/20 active:scale-95 transition-all"><Trash2 className="w-3.5 h-3.5"/></button>
                       </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}