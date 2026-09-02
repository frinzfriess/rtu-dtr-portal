import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import DtrForm from './DtrForm';
import { LogOut, Clock, CalendarDays, Building2, GraduationCap, User, Trash2, Edit, X, Save, FileSpreadsheet, Search, SearchX, CheckCircle2 } from 'lucide-react';
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
    if(!window.confirm("Are you sure you want to delete this record?")) return;
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
    } catch (err) { alert("Failed to update record."); } 
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
      headerRow.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; });

      const sortedDtr = [...dtr].sort((a,b) => new Date(a.date) - new Date(b.date));
      sortedDtr.forEach(entry => {
          const row = worksheet.addRow([ new Date(entry.date).toLocaleDateString(), entry.time_in.substring(0, 5), entry.time_out.substring(0, 5), entry.hours_rendered, entry.remarks || '' ]);
          row.eachCell(cell => { cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} }; cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' }; });
      });
      worksheet.addRow([]);
      const totalRow = worksheet.addRow(['', '', 'TOTAL HOURS:', totalHours.toFixed(2), '']);
      totalRow.font = { bold: true, size: 12 }; totalRow.getCell(4).font = { bold: true, size: 12, color: { argb: 'FF3B82F6' } };
      
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `DTR_${profile.full_name || 'Student'}.xlsx`);
    } catch (err) { alert("Excel export failed."); }
  };

  if (loading || !profile) return null;
  const progressPercentage = Math.min((totalHours / profile.required_hours) * 100, 100);
  const remainingHours = Math.max(profile.required_hours - totalHours, 0);
  const filteredDtr = dtr.filter(e => e.remarks?.toLowerCase().includes(searchTerm.toLowerCase()) || e.date.includes(searchTerm));

  return (
    <div className="max-w-7xl mx-auto py-8 px-4 sm:px-6 lg:px-8 relative font-sans text-slate-100 bg-dark-950 min-h-screen">
      
      {/* Edit Modal */}
      {editingEntry && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4 z-50">
          <div className="bg-dark-900 border border-slate-800 rounded-3xl p-8 w-full max-w-md shadow-2xl relative">
            <button onClick={() => setEditingEntry(null)} className="absolute top-6 right-6 text-slate-400 hover:text-white transition-colors bg-slate-800 p-2 rounded-full">
              <X className="w-5 h-5" />
            </button>
            <h3 className="text-2xl font-black text-white mb-6 flex items-center gap-2"><Edit className="w-5 h-5 text-blue-400"/> Edit Record</h3>
            
            <div className="space-y-4 text-sm">
              <div>
                <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-4 py-3 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 outline-none font-bold text-white" />
              </div>
              <div className="flex gap-4">
                <div className="flex-1">
                  <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-4 py-3 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 outline-none font-bold text-white" />
                </div>
                <div className="flex-1">
                  <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-4 py-3 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 outline-none font-bold text-white" />
                </div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-4 py-3 bg-dark-950 border border-slate-800 rounded-xl focus:border-blue-500 outline-none font-bold text-white" />
              </div>
              <label className="flex items-center space-x-2 cursor-pointer mt-2 bg-dark-950 p-3 rounded-xl border border-slate-800">
                <input type="checkbox" checked={deductLunchEdit} onChange={(e) => setDeductLunchEdit(e.target.checked)} className="w-4 h-4 text-blue-500 rounded border-slate-700 bg-dark-900" />
                <span className="font-bold text-slate-300">Auto-deduct 1hr lunch</span>
              </label>
            </div>

            <div className="flex gap-3 mt-8">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold py-3 rounded-xl transition-colors">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-1 bg-blue-600 hover:bg-blue-500 text-white font-black py-3 rounded-xl shadow-glow transition-all flex items-center justify-center gap-2">
                <Save className="w-4 h-4"/> {isSavingEdit ? 'Saving...' : 'Save Changes'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Top Header */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4 border-b border-slate-800 pb-6">
        <div>
          <p className="text-blue-400 font-bold text-xs tracking-wider uppercase mb-1">{getGreeting()}</p>
          <h1 className="text-3xl font-black text-white tracking-tight">{profile.full_name || session.user.email.split('@')[0]}</h1>
        </div>
        <div className="flex gap-3 w-full md:w-auto text-sm">
          <button onClick={exportStyledExcel} className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-blue-600 text-white hover:bg-blue-500 shadow-lg shadow-blue-600/25 font-bold px-5 py-3 rounded-xl transition-all">
            <FileSpreadsheet className="w-4 h-4" /> Export DTR
          </button>
          <button onClick={() => supabase.auth.signOut()} className="flex-none flex items-center justify-center gap-2 bg-dark-900 text-slate-400 hover:text-rose-400 hover:bg-rose-500/10 font-bold px-5 py-3 rounded-xl border border-slate-800 transition-all">
            <LogOut className="w-4 h-4" /> Sign Out
          </button>
        </div>
      </div>
      
      {/* Profile & Stats */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 mb-8">
        
        <div className="lg:col-span-2 bg-dark-900 border border-slate-800 p-6 sm:p-8 rounded-3xl shadow-soft">
          <h2 className="text-lg font-black text-white mb-5 flex items-center gap-2"><User className="w-5 h-5 text-blue-400"/> Student Profile</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
            <div className="bg-dark-950 p-4 rounded-2xl border border-slate-800 focus-within:border-blue-500 transition-all">
              <label className="text-[11px] uppercase tracking-wider text-slate-400 font-bold flex items-center gap-1.5 mb-1"><User className="w-3 h-3 text-blue-400"/> Full Name</label>
              <input type="text" defaultValue={profile.full_name || ''} placeholder="Juan Dela Cruz" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-white" />
            </div>
            <div className="bg-dark-950 p-4 rounded-2xl border border-slate-800 focus-within:border-blue-500 transition-all">
              <label className="text-[11px] uppercase tracking-wider text-slate-400 font-bold flex items-center gap-1.5 mb-1"><Building2 className="w-3 h-3 text-blue-400"/> Company</label>
              <input type="text" defaultValue={profile.company_name} placeholder="Tech Corp Inc." onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-white" />
            </div>
            <div className="bg-dark-950 p-4 rounded-2xl border border-slate-800 focus-within:border-blue-500 transition-all sm:col-span-2">
              <label className="text-[11px] uppercase tracking-wider text-slate-400 font-bold flex items-center gap-1.5 mb-1"><GraduationCap className="w-3 h-3 text-blue-400"/> School / University</label>
              <input type="text" defaultValue={profile.school} onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-white" />
            </div>
          </div>
        </div>

        <div className="bg-dark-900 border border-slate-800 p-6 sm:p-8 rounded-3xl shadow-soft text-white flex flex-col justify-between relative overflow-hidden">
          <div className="absolute top-0 right-0 p-8 opacity-10 pointer-events-none"><Clock className="w-32 h-32 text-blue-400"/></div>
          <div className="relative z-10">
            <div className="flex justify-between items-start mb-6">
              <div>
                <p className="text-blue-400 font-bold mb-1 text-xs uppercase tracking-wider">Total Rendered</p>
                <h3 className="text-4xl font-black tracking-tight">{totalHours.toFixed(2)}<span className="text-xl">h</span></h3>
              </div>
              <div className="bg-dark-950 border border-slate-800 px-3 py-2 rounded-xl flex items-center gap-2">
                <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-12 bg-transparent outline-none font-bold text-right text-blue-400 text-sm" />
                <span className="text-[10px] font-bold text-slate-500 uppercase">goal</span>
              </div>
            </div>
            <div className="space-y-2">
              <div className="flex justify-between text-xs font-bold text-slate-400">
                <span>Progress</span>
                <span className="text-blue-400">{progressPercentage.toFixed(1)}%</span>
              </div>
              <div className="w-full bg-dark-950 rounded-full h-2.5 overflow-hidden border border-slate-800">
                <div className="bg-gradient-to-r from-blue-500 to-indigo-500 h-full rounded-full transition-all duration-1000 shadow-glow" style={{ width: `${progressPercentage}%` }}></div>
              </div>
            </div>
            <div className="mt-6 pt-5 border-t border-slate-800 grid grid-cols-2 gap-4">
              <div>
                <p className="text-slate-500 text-[11px] font-bold mb-0.5 uppercase tracking-wider">Days Present</p>
                <p className="font-black text-lg flex items-center gap-1.5 text-white"><CalendarDays className="w-4 h-4 text-blue-400"/> {dtr.length}</p>
              </div>
              <div>
                <p className="text-slate-500 text-[11px] font-bold mb-0.5 uppercase tracking-wider">Avg Hrs/Day</p>
                <p className="font-black text-lg flex items-center gap-1.5 text-white"><Clock className="w-4 h-4 text-blue-400"/> {dtr.length > 0 ? (totalHours / dtr.length).toFixed(1) : 0}</p>
              </div>
            </div>
          </div>
        </div>
      </div>

      <DtrForm userId={session.user.id} onEntryAdded={fetchDTR} />

      {/* Table Section */}
      <div className="bg-dark-900 border border-slate-800 rounded-3xl shadow-soft overflow-hidden mt-8">
        <div className="px-6 py-5 border-b border-slate-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-dark-900/50">
          <h3 className="font-black text-white flex items-center gap-2 text-base"><Clock className="w-5 h-5 text-blue-400"/> Timesheet History</h3>
          
          <div className="relative w-full sm:w-72">
             <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-500" /></div>
             <input type="text" placeholder="Search records..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2 bg-dark-950 border border-slate-800 rounded-xl text-xs font-bold text-white outline-none focus:border-blue-500" />
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse whitespace-nowrap text-xs">
            <thead>
              <tr className="bg-dark-950 border-b border-slate-800 text-[11px] text-slate-400 uppercase tracking-wider">
                <th className="p-4 px-6">Date</th>
                <th className="p-4 px-6">Time In</th>
                <th className="p-4 px-6">Time Out</th>
                <th className="p-4 px-6">Hours</th>
                <th className="p-4 px-6">Remarks</th>
                <th className="p-4 px-6 text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/50">
              {filteredDtr.length === 0 ? (
                <tr><td colSpan="6" className="p-12 text-center text-slate-500">Your timesheet is empty. Add your first record above!</td></tr>
              ) : (
                filteredDtr.map((entry) => (
                  <tr key={entry.id} className="hover:bg-dark-950/40 transition-colors group">
                    <td className="p-4 px-6 font-bold text-slate-300">{formatDateString(entry.date)}</td>
                    <td className="p-4 px-6 text-slate-400">{entry.time_in.substring(0, 5)}</td>
                    <td className="p-4 px-6 text-slate-400">{entry.time_out.substring(0, 5)}</td>
                    <td className="p-4 px-6"><span className="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-bold bg-blue-500/10 text-blue-400 border border-blue-500/20">{entry.hours_rendered}h</span></td>
                    <td className="p-4 px-6 text-slate-300 max-w-[250px] truncate">{entry.remarks || '-'}</td>
                    <td className="p-4 px-6 text-right">
                      <div className="flex justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button onClick={() => handleEditClick(entry)} className="text-blue-400 hover:text-blue-300 bg-dark-950 hover:bg-slate-800 p-2 rounded-lg transition-all border border-slate-800" title="Edit">
                          <Edit className="w-4 h-4"/>
                        </button>
                        <button onClick={() => deleteEntry(entry.id)} className="text-rose-400 hover:text-rose-300 bg-dark-950 hover:bg-rose-500/10 p-2 rounded-lg transition-all border border-slate-800" title="Delete">
                          <Trash2 className="w-4 h-4"/>
                        </button>
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
  );
}