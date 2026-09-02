import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import { 
  LogOut, Clock, CalendarDays, User, 
  Trash2, Edit, X, Save, FileSpreadsheet, Search, Calculator, 
  CheckCircle2, Terminal, UtensilsCrossed, LayoutDashboard, 
  FileText, UserCheck, ChevronRight, PlusCircle, Shield, Sun, Moon,
  Timer, BarChart3, TrendingUp, Award, Zap, History
} from 'lucide-react';
import { calculateHoursRendered, getGreeting, formatDateString, getCurrentTime, getInitials } from '../utils';

export default function Dashboard({ session, darkMode, setDarkMode }) {
  const [profile, setProfile] = useState(null);
  const [dtr, setDtr] = useState([]);
  const [totalHours, setTotalHours] = useState(0);
  const [loading, setLoading] = useState(true);
  const [searchTerm, setSearchTerm] = useState('');
  const [activeTab, setActiveTab] = useState('overview');

  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [timeIn, setTimeIn] = useState('08:00');
  const [timeOut, setTimeOut] = useState('17:00');
  const [remarks, setRemarks] = useState('');
  const [deductLunch, setDeductLunch] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [successStatus, setSuccessStatus] = useState(false);

  const [isTimerRunning, setIsTimerRunning] = useState(false);
  const [timerSeconds, setTimerSeconds] = useState(0);

  const [editingEntry, setEditingEntry] = useState(null);
  const [editForm, setEditForm] = useState({});
  const [isSavingEdit, setIsSavingEdit] = useState(false);
  const [deductLunchEdit, setDeductLunchEdit] = useState(true);

  const [targetDaysPerWeek, setTargetDaysPerWeek] = useState(5);
  const [targetHoursPerDay, setTargetHoursPerDay] = useState(8);

  useEffect(() => { fetchProfile(); fetchDTR(); }, [session]);

  useEffect(() => {
    let interval = null;
    if (isTimerRunning) {
      interval = setInterval(() => setTimerSeconds(prev => prev + 1), 1000);
    } else {
      clearInterval(interval);
    }
    return () => clearInterval(interval);
  }, [isTimerRunning]);

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

  const handleLogSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);
    const hours_rendered = calculateHoursRendered(timeIn, timeOut, deductLunch);

    if (hours_rendered <= 0) {
      alert("Invalid time range. Time out must be after Time in.");
      setSubmitting(false);
      return;
    }

    try {
      const { error } = await supabase.from('dtr_entries').insert([{
        user_id: session.user.id, date, time_in: `${timeIn}:00`, time_out: `${timeOut}:00`, hours_rendered, remarks
      }]);
      if (error) throw error;
      
      setRemarks('');
      fetchDTR();
      setSuccessStatus(true);
      setTimeout(() => setSuccessStatus(false), 2500);
    } catch (error) { alert("Failed to add entry. " + error.message); } 
    finally { setSubmitting(false); }
  };

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
      const headerRow = worksheet.addRow(['Date', 'Time In', 'Time Out', 'Hours', 'Remarks/Tasks']);
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
      saveAs(new Blob([buffer]), `DTR_${profile.full_name || 'Student'}.xlsx`);
    } catch (err) { alert("Failed to export Excel."); }
  };

  if (loading || !profile) return null;
  const progressPercentage = Math.min((totalHours / profile.required_hours) * 100, 100);
  const remainingHours = Math.max(profile.required_hours - totalHours, 0);
  const filteredDtr = dtr.filter(e => e.remarks?.toLowerCase().includes(searchTerm.toLowerCase()) || e.date.includes(searchTerm));

  const hoursPerWeek = targetDaysPerWeek * targetHoursPerDay;
  const weeksRemaining = hoursPerWeek > 0 ? remainingHours / hoursPerWeek : 0;
  const estimatedCompletionDays = Math.ceil(weeksRemaining * 7);
  const estimatedCompletionDate = new Date();
  estimatedCompletionDate.setDate(estimatedCompletionDate.getDate() + estimatedCompletionDays);

  const averageHoursPerDay = dtr.length > 0 ? (totalHours / dtr.length).toFixed(2) : 0;
  const efficiencyRate = profile.required_hours > 0 ? ((totalHours / profile.required_hours) * 100).toFixed(1) : 0;

  return (
    <div className="min-h-screen flex flex-col md:flex-row bg-slate-100 dark:bg-slate-950 text-slate-900 dark:text-slate-100 font-sans transition-colors duration-300 pb-28 md:pb-0 overflow-x-hidden">
      
      {/* EDIT MODAL */}
      {editingEntry && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4 z-50">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 w-full max-w-md shadow-2xl relative font-mono">
            <button onClick={() => setEditingEntry(null)} className="absolute top-6 right-6 text-slate-400 hover:text-slate-900 dark:hover:text-white bg-slate-100 dark:bg-slate-800 p-2 rounded-full transition-colors">
              <X className="w-5 h-5" />
            </button>
            <h3 className="text-xl font-black text-slate-900 dark:text-white mb-6 flex items-center gap-2"><Edit className="w-5 h-5 text-cyan-500"/> Edit Record</h3>
            
            <div className="space-y-4 text-xs">
              <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full bg-transparent border-none p-0 m-0 outline-none text-slate-900 dark:text-white font-bold block appearance-none" />
              </div>
              <div className="flex gap-4">
                <div className="flex-1 bg-slate-50 dark:bg-slate-950 px-4 py-3 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full bg-transparent border-none p-0 m-0 outline-none text-slate-900 dark:text-white font-bold block appearance-none" />
                </div>
                <div className="flex-1 bg-slate-50 dark:bg-slate-950 px-4 py-3 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full bg-transparent border-none p-0 m-0 outline-none text-slate-900 dark:text-white font-bold block appearance-none" />
                </div>
              </div>
              <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full bg-transparent border-none p-0 m-0 outline-none text-slate-900 dark:text-white font-bold block appearance-none" />
              </div>
            </div>

            <div className="flex gap-3 mt-8 font-mono text-xs">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-200 dark:bg-slate-800 text-slate-800 dark:text-slate-300 font-bold py-3.5 rounded-xl hover:bg-slate-300 dark:hover:bg-slate-700 transition-colors">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-1 bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black py-3.5 rounded-xl flex items-center justify-center gap-2 transition-colors">
                <Save className="w-4 h-4"/> {isSavingEdit ? 'Saving...' : 'Save Changes'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DESKTOP SIDEBAR */}
      <aside className="hidden md:flex w-72 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 flex-col justify-between p-6 flex-shrink-0">
        <div>
          <div className="flex items-center gap-3 pb-6 mb-6 border-b border-slate-200 dark:border-slate-800">
            <div className="bg-cyan-500/10 p-3 rounded-2xl text-cyan-600 dark:text-cyan-400 border border-cyan-500/20"><Shield className="w-6 h-6" /></div>
            <div>
              <span className="text-[10px] font-mono font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest block">RTU Portal</span>
              <h1 className="text-base font-black text-slate-900 dark:text-white font-mono">DTR SYSTEM</h1>
            </div>
          </div>

          <div className="flex items-center gap-3 mb-6 p-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-2xl">
            <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-indigo-500 to-cyan-500 flex items-center justify-center text-white font-black text-sm shadow-md">
              {getInitials(profile?.full_name)}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-bold text-slate-900 dark:text-white truncate">{profile?.full_name || 'Student'}</p>
              <p className="text-[10px] text-slate-500 truncate">{profile?.company_name || 'No Company'}</p>
            </div>
          </div>

          <nav className="space-y-2 font-mono text-xs">
            <button onClick={() => setActiveTab('overview')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'overview' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><LayoutDashboard className="w-4 h-4"/> Overview</div>
              <ChevronRight className="w-3.5 h-3.5 opacity-60"/>
            </button>
            <button onClick={() => setActiveTab('log')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'log' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><PlusCircle className="w-4 h-4"/> Log Hours</div>
              <ChevronRight className="w-3.5 h-3.5 opacity-60"/>
            </button>
            <button onClick={() => setActiveTab('analytics')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'analytics' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><BarChart3 className="w-4 h-4"/> Analytics</div>
              <ChevronRight className="w-3.5 h-3.5 opacity-60"/>
            </button>
            <button onClick={() => setActiveTab('history')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'history' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><History className="w-4 h-4"/> History</div>
              <span className="bg-slate-200 dark:bg-slate-800 px-2 py-0.5 rounded-md font-bold">{dtr.length}</span>
            </button>
            <button onClick={() => setActiveTab('estimator')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'estimator' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><Calculator className="w-4 h-4"/> Estimator</div>
              <ChevronRight className="w-3.5 h-3.5 opacity-60"/>
            </button>
            <button onClick={() => setActiveTab('profile')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl transition-all ${activeTab === 'profile' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><UserCheck className="w-4 h-4"/> Profile</div>
              <ChevronRight className="w-3.5 h-3.5 opacity-60"/>
            </button>
          </nav>
        </div>

        <div className="pt-6 border-t border-slate-200 dark:border-slate-800 space-y-4">
          <div className="flex gap-2 font-mono">
            <button onClick={exportStyledExcel} className="flex-1 flex items-center justify-center gap-1.5 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-600 dark:text-emerald-400 border border-emerald-500/30 py-2.5 rounded-xl text-xs font-bold transition-colors">
              <FileSpreadsheet className="w-3.5 h-3.5"/> Export
            </button>
            <button onClick={() => setDarkMode(!darkMode)} className="p-2.5 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors">
              {darkMode ? <Sun className="w-4 h-4 text-amber-400"/> : <Moon className="w-4 h-4"/>}
            </button>
            <button onClick={() => supabase.auth.signOut()} className="p-2.5 rounded-xl bg-rose-500/10 hover:bg-rose-500/20 text-rose-600 border border-rose-500/30 transition-colors">
              <LogOut className="w-3.5 h-3.5"/>
            </button>
          </div>
        </div>
      </aside>

      {/* MOBILE BOTTOM NAV */}
      <nav aria-label="Mobile Navigation" className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-white/95 dark:bg-slate-900/95 backdrop-blur-xl border-t border-slate-200 dark:border-slate-800 px-2 py-2 flex justify-around items-center shadow-[0_-10px_40px_-15px_rgba(0,0,0,0.1)]">
        <button onClick={() => setActiveTab('overview')} className={`flex flex-col items-center p-1.5 rounded-xl transition-all ${activeTab === 'overview' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <LayoutDashboard className="w-5 h-5 mb-0.5"/><span className="text-[9px] font-mono">Home</span>
        </button>
        <button onClick={() => setActiveTab('log')} className={`flex flex-col items-center p-1.5 rounded-xl transition-all ${activeTab === 'log' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <PlusCircle className="w-5 h-5 mb-0.5"/><span className="text-[9px] font-mono">Log</span>
        </button>
        <button onClick={() => setActiveTab('analytics')} className={`flex flex-col items-center p-1.5 rounded-xl transition-all ${activeTab === 'analytics' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <BarChart3 className="w-5 h-5 mb-0.5"/><span className="text-[9px] font-mono">Stats</span>
        </button>
        <button onClick={() => setActiveTab('history')} className={`flex flex-col items-center p-1.5 rounded-xl transition-all ${activeTab === 'history' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <History className="w-5 h-5 mb-0.5"/><span className="text-[9px] font-mono">History</span>
        </button>
        <button onClick={() => setActiveTab('estimator')} className={`flex flex-col items-center p-1.5 rounded-xl transition-all ${activeTab === 'estimator' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <Calculator className="w-5 h-5 mb-0.5"/><span className="text-[9px] font-mono">Estimator</span>
        </button>
      </nav>

      {/* MAIN CONTAINER */}
      <main className="flex-1 flex flex-col min-w-0 w-full overflow-y-auto bg-slate-100 dark:bg-slate-950 relative">
        
        <header className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-md border-b border-slate-200 dark:border-slate-800 px-4 sm:px-6 py-4 flex items-center justify-between sticky top-0 z-30">
          <div className="min-w-0 pr-2">
            <div className="flex items-center gap-2 text-cyan-600 dark:text-cyan-400 text-[11px] font-mono font-bold uppercase tracking-widest">
              <span className="w-2 h-2 rounded-full bg-cyan-500 animate-pulse"></span>
              <span>{getGreeting()}</span>
            </div>
            <h2 className="text-base sm:text-xl font-black text-slate-900 dark:text-white truncate">
              {activeTab === 'overview' && 'Dashboard Overview'}
              {activeTab === 'log' && 'Log Daily Hours'}
              {activeTab === 'analytics' && 'Performance Analytics'}
              {activeTab === 'history' && 'Timesheet History'}
              {activeTab === 'estimator' && 'OJT Estimator'}
              {activeTab === 'profile' && 'Student Profile'}
            </h2>
          </div>

          <div className="flex items-center gap-2 font-mono text-xs">
            <button onClick={() => setDarkMode(!darkMode)} className="md:hidden p-2 rounded-xl bg-slate-200 dark:bg-slate-800 text-slate-800 dark:text-slate-200">
              {darkMode ? <Sun className="w-4 h-4 text-amber-400"/> : <Moon className="w-4 h-4"/>}
            </button>
            <button onClick={() => setActiveTab('profile')} className="md:hidden p-2 rounded-xl bg-cyan-500/10 text-cyan-600 border border-cyan-500/20">
              <User className="w-4 h-4"/>
            </button>
            <button onClick={exportStyledExcel} className="hidden sm:flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-500 text-white font-bold px-4 py-2.5 rounded-xl shadow-md transition-all">
              <FileSpreadsheet className="w-4 h-4"/> Export DTR
            </button>
          </div>
        </header>

        <div className="p-4 sm:p-6 lg:p-8 max-w-6xl w-full mx-auto space-y-6 pb-24 md:pb-8">

          {/* OVERVIEW TAB */}
          {activeTab === 'overview' && (
            <div className="space-y-6 w-full animate-fadeIn">
              <div className="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 text-white rounded-3xl p-6 sm:p-8 shadow-xl relative overflow-hidden flex flex-col justify-between border border-slate-800">
                <div className="relative z-10">
                  <div className="flex items-center gap-2 mb-4">
                    <span className="bg-white/10 border border-white/20 px-3 py-1.5 rounded-full text-[10px] font-mono font-bold uppercase backdrop-blur-md flex items-center gap-1.5">
                       <Zap className="w-3 h-3 text-cyan-400"/> RTU OJT Portal
                    </span>
                  </div>
                  <h3 className="text-2xl sm:text-4xl font-black tracking-tight mb-2 text-transparent bg-clip-text bg-gradient-to-r from-white to-slate-400">Welcome, {profile.full_name?.split(' ')[0] || 'Student'}!</h3>
                  <p className="text-xs sm:text-sm text-slate-300 max-w-xl mb-8 leading-relaxed">
                    You have successfully completed <strong className="text-cyan-400">{progressPercentage.toFixed(1)}%</strong> of your required internship hours at <strong className="text-white">{profile.company_name}</strong>.
                  </p>
                </div>
                
                <div className="relative z-10 w-full bg-white/5 border border-white/10 rounded-2xl p-4 sm:p-5 backdrop-blur-sm">
                  <div className="flex justify-between items-end mb-2">
                    <span className="text-[10px] font-mono text-slate-400 uppercase tracking-widest">Progress</span>
                    <span className="text-sm font-black text-cyan-400 font-mono">{totalHours.toFixed(1)} / {profile.required_hours} hrs</span>
                  </div>
                  <div className="w-full bg-slate-950 rounded-full h-3 overflow-hidden border border-slate-800 shadow-inner">
                    <div className="bg-gradient-to-r from-cyan-500 to-indigo-500 h-full rounded-full relative" style={{ width: `${progressPercentage}%` }}></div>
                  </div>
                </div>
              </div>

              {/* LIVE TIMER */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 sm:p-6 rounded-3xl shadow-soft flex flex-col sm:flex-row items-center justify-between gap-5 w-full transition-colors">
                <div className="flex items-center gap-4">
                  <div className="bg-cyan-500/10 text-cyan-600 dark:text-cyan-400 p-3.5 rounded-2xl border border-cyan-500/20">
                    <Timer className="w-7 h-7 animate-pulse"/>
                  </div>
                  <div>
                    <h4 className="text-[11px] font-mono font-bold text-slate-400 uppercase tracking-widest mb-0.5">Live Shift Timer</h4>
                    <p className="text-3xl font-black font-mono text-slate-900 dark:text-white tracking-tight">
                      {Math.floor(timerSeconds / 3600).toString().padStart(2, '0')}:
                      {Math.floor((timerSeconds % 3600) / 60).toString().padStart(2, '0')}:
                      {(timerSeconds % 60).toString().padStart(2, '0')}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2.5 font-mono text-xs w-full sm:w-auto">
                  {!isTimerRunning ? (
                    <button onClick={() => setIsTimerRunning(true)} className="flex-1 sm:flex-none bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-black px-6 py-3.5 rounded-xl shadow-neon transition-all">Start Shift</button>
                  ) : (
                    <button onClick={() => setIsTimerRunning(false)} className="flex-1 sm:flex-none bg-rose-500 hover:bg-rose-400 text-white font-black px-6 py-3.5 rounded-xl shadow-lg transition-all">Pause</button>
                  )}
                  <button onClick={() => { setIsTimerRunning(false); setTimerSeconds(0); }} className="bg-slate-100 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 text-slate-700 dark:text-slate-300 font-bold px-4 py-3.5 rounded-xl transition-all">Reset</button>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 w-full">
                <div className="lg:col-span-1 space-y-4">
                  <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft flex items-center justify-between">
                    <div>
                      <span className="text-[10px] font-mono font-bold text-slate-400 uppercase">Days Logged</span>
                      <h3 className="text-2xl font-black font-mono mt-1 text-slate-900 dark:text-white">{dtr.length} <span className="text-xs text-slate-400 font-sans font-normal">shifts</span></h3>
                    </div>
                    <div className="bg-indigo-500/10 text-indigo-500 p-3 rounded-2xl"><CalendarDays className="w-5 h-5"/></div>
                  </div>
                  <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft flex items-center justify-between">
                    <div>
                      <span className="text-[10px] font-mono font-bold text-slate-400 uppercase">Remaining</span>
                      <h3 className="text-2xl font-black font-mono mt-1 text-emerald-500">{remainingHours.toFixed(1)} <span className="text-xs text-slate-400 font-sans font-normal">hrs</span></h3>
                    </div>
                    <div className="bg-emerald-500/10 text-emerald-500 p-3 rounded-2xl"><Award className="w-5 h-5"/></div>
                  </div>
                </div>

                <div className="lg:col-span-2 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-6 rounded-3xl shadow-soft flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-sm font-black text-slate-900 dark:text-white font-mono uppercase tracking-wider flex items-center gap-2">
                       <Clock className="w-4 h-4 text-cyan-500"/> Recent Activity
                    </h3>
                    <button onClick={() => setActiveTab('history')} className="text-[10px] font-mono font-bold text-cyan-600 dark:text-cyan-400 hover:underline">View All &rarr;</button>
                  </div>
                  
                  {dtr.length === 0 ? (
                    <div className="flex-1 flex flex-col items-center justify-center py-8 text-slate-400 bg-slate-50 dark:bg-slate-950/50 rounded-2xl border border-dashed border-slate-200 dark:border-slate-800">
                      <History className="w-8 h-8 mb-2 opacity-50"/>
                      <span className="text-xs font-mono">No recent logs.</span>
                    </div>
                  ) : (
                    <div className="space-y-3 font-mono">
                      {dtr.slice(0, 3).map((entry) => (
                        <div key={entry.id} className="flex items-center justify-between p-3.5 bg-slate-50 dark:bg-slate-950 rounded-2xl border border-slate-200 dark:border-slate-800">
                          <div className="flex items-center gap-3">
                            <div className="w-2 h-2 rounded-full bg-emerald-500"></div>
                            <div>
                              <p className="text-xs font-bold text-slate-900 dark:text-white">{formatDateString(entry.date)}</p>
                              <p className="text-[10px] text-slate-500">{entry.time_in.substring(0, 5)} &rarr; {entry.time_out.substring(0, 5)}</p>
                            </div>
                          </div>
                          <span className="px-3 py-1 rounded-lg text-xs font-bold bg-cyan-500/10 text-cyan-700 dark:text-cyan-400 border border-cyan-500/20">
                            +{entry.hours_rendered}h
                          </span>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* SAFARI-FIXED LOG HOURS TAB */}
          {activeTab === 'log' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-5 sm:p-8 shadow-soft max-w-xl mx-auto w-full box-border animate-fadeIn">
              
              <div className="bg-gradient-to-br from-cyan-500/10 to-indigo-500/10 p-4 rounded-2xl border border-cyan-500/20 mb-6 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="bg-cyan-500 text-slate-950 p-2.5 rounded-xl font-bold shadow-glow">
                    <PlusCircle className="w-5 h-5"/>
                  </div>
                  <div>
                    <h3 className="text-sm font-black text-slate-900 dark:text-white font-mono uppercase tracking-wide">New Shift Entry</h3>
                    <p className="text-[10px] sm:text-[11px] text-slate-500 dark:text-slate-400 font-mono">Record your daily hours.</p>
                  </div>
                </div>
              </div>

              <form onSubmit={handleLogSubmit} className="space-y-4 w-full block">
                <div className="w-full">
                  <button
                    type="button"
                    onClick={() => setDeductLunch(!deductLunch)}
                    className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl border transition-all select-none ${
                      deductLunch ? 'bg-cyan-500/10 border-cyan-500/40 text-cyan-700 dark:text-cyan-300' : 'bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800 text-slate-400'
                    }`}
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <UtensilsCrossed className="w-4 h-4 flex-shrink-0 text-cyan-500" />
                      <span className="text-xs font-bold font-mono truncate">Auto-deduct 1h lunch break</span>
                    </div>
                    <div className={`w-5 h-5 rounded-full flex items-center justify-center border flex-shrink-0 transition-colors ${
                      deductLunch ? 'bg-cyan-500 border-cyan-600' : 'bg-transparent border-slate-400'
                    }`}>
                      {deductLunch && <div className="w-2 h-2 rounded-full bg-slate-950"></div>}
                    </div>
                  </button>
                </div>

                <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 focus-within:ring-2 focus-within:ring-cyan-500/25 transition-all shadow-sm">
                  <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Date</label>
                  <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
                    className="w-full bg-transparent border-none p-0 m-0 text-sm font-bold text-slate-900 dark:text-white outline-none focus:ring-0 font-mono block appearance-none" />
                </div>

                <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 focus-within:ring-2 focus-within:ring-cyan-500/25 transition-all shadow-sm">
                  <div className="flex justify-between items-center mb-1.5 w-full">
                     <label className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time In</label>
                     <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 hover:bg-cyan-500/20 px-2.5 py-1 rounded-md font-mono flex items-center gap-1 transition-colors">
                       <Clock className="w-3 h-3"/> Now
                     </button>
                  </div>
                  <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
                    className="w-full bg-transparent border-none p-0 m-0 text-lg sm:text-xl font-black text-slate-900 dark:text-white outline-none focus:ring-0 font-mono block appearance-none" />
                </div>

                <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 focus-within:ring-2 focus-within:ring-cyan-500/25 transition-all shadow-sm">
                  <div className="flex justify-between items-center mb-1.5 w-full">
                     <label className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time Out</label>
                     <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 hover:bg-cyan-500/20 px-2.5 py-1 rounded-md font-mono flex items-center gap-1 transition-colors">
                       <Clock className="w-3 h-3"/> Now
                     </button>
                  </div>
                  <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
                    className="w-full bg-transparent border-none p-0 m-0 text-lg sm:text-xl font-black text-slate-900 dark:text-white outline-none focus:ring-0 font-mono block appearance-none" />
                </div>

                <div className="w-full bg-slate-50 dark:bg-slate-950 px-4 py-3.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 focus-within:ring-2 focus-within:ring-cyan-500/25 transition-all shadow-sm">
                  <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Remarks / Tasks Accomplished</label>
                  <textarea rows="3" placeholder="E.g., System debugging, UI updates, documentation..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
                    className="w-full bg-transparent border-none p-0 m-0 text-xs sm:text-sm font-bold text-slate-900 dark:text-white outline-none focus:ring-0 font-sans resize-none block appearance-none" />
                </div>

                <div className="pt-4 w-full">
                  <button type="submit" disabled={submitting} 
                    className={`w-full font-black py-4 rounded-2xl transition-all shadow-glow font-mono text-xs tracking-wider uppercase flex items-center justify-center gap-2 block box-border ${
                      successStatus ? 'bg-emerald-500 text-white shadow-neon scale-[0.98]' : 'bg-gradient-to-r from-cyan-500 to-indigo-600 hover:from-cyan-400 text-slate-950 active:scale-[0.98]'
                    }`}>
                    {successStatus ? <><CheckCircle2 className="w-4 h-4"/> Record Logged Successfully!</> : 'Submit Daily Shift Record'}
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* ANALYTICS TAB */}
          {activeTab === 'analytics' && (
            <div className="space-y-6 w-full animate-fadeIn">
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-soft">
                <div className="flex items-center gap-3 mb-8 pb-4 border-b border-slate-100 dark:border-slate-800">
                  <div className="bg-cyan-500/10 p-3 rounded-2xl text-cyan-600 dark:text-cyan-400 border border-cyan-500/30">
                    <BarChart3 className="w-6 h-6"/>
                  </div>
                  <div>
                    <h3 className="text-lg font-black text-slate-900 dark:text-white uppercase tracking-wider font-mono">Performance Analytics</h3>
                    <p className="text-xs text-slate-500 dark:text-slate-400 font-mono">Detailed breakdown of your internship statistics.</p>
                  </div>
                </div>

                <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 font-mono">
                  <div className="bg-slate-50 dark:bg-slate-950 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
                    <span className="text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest block mb-2">Average / Shift</span>
                    <p className="text-xl sm:text-2xl font-black text-slate-900 dark:text-white">{averageHoursPerDay} <span className="text-[10px] sm:text-xs font-bold text-cyan-500">hrs</span></p>
                  </div>

                  <div className="bg-slate-50 dark:bg-slate-950 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
                    <span className="text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest block mb-2">Total Shifts</span>
                    <p className="text-xl sm:text-2xl font-black text-slate-900 dark:text-white">{dtr.length} <span className="text-[10px] sm:text-xs font-bold text-indigo-500">days</span></p>
                  </div>

                  <div className="bg-slate-50 dark:bg-slate-950 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
                    <span className="text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest block mb-2">Efficiency</span>
                    <p className="text-xl sm:text-2xl font-black text-cyan-600 dark:text-cyan-400">{efficiencyRate}%</p>
                  </div>

                  <div className="bg-slate-50 dark:bg-slate-950 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm">
                    <span className="text-[9px] sm:text-[10px] text-slate-400 uppercase tracking-widest block mb-2">Completion Date</span>
                    <p className="text-sm sm:text-base font-black text-emerald-600 dark:text-emerald-400 truncate mt-1">{remainingHours <= 0 ? 'Done! 🎉' : estimatedCompletionDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}</p>
                  </div>
                </div>

                <div className="mt-6 bg-gradient-to-r from-slate-900 to-indigo-950 text-white p-6 rounded-2xl border border-slate-800 flex flex-col sm:flex-row items-center gap-5 shadow-lg">
                  <div className="bg-cyan-500/20 text-cyan-400 p-4 rounded-full border border-cyan-500/30">
                     <TrendingUp className="w-8 h-8"/>
                  </div>
                  <div className="text-center sm:text-left">
                    <h4 className="text-sm sm:text-base font-black text-white font-mono uppercase tracking-wide">Velocity Insight</h4>
                    <p className="text-xs sm:text-sm text-slate-300 font-sans mt-1.5 leading-relaxed">
                      You are logging an average of <strong>{averageHoursPerDay} hours</strong> per shift. 
                      At this current pace, you are on track to complete your remaining <strong>{remainingHours.toFixed(1)} hours</strong> smoothly. Keep it up!
                    </p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* HISTORY TAB */}
          {activeTab === 'history' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl shadow-soft overflow-hidden w-full animate-fadeIn">
              <div className="px-4 sm:px-6 py-5 border-b border-slate-100 dark:border-slate-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-slate-50 dark:bg-slate-950/50">
                <h3 className="font-black text-slate-900 dark:text-white flex items-center gap-2 text-base font-mono uppercase"><History className="w-5 h-5 text-cyan-500"/> Timesheet History</h3>
                <div className="relative w-full sm:w-72">
                   <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-400" /></div>
                   <input type="text" placeholder="Search records..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2.5 bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 focus:ring-2 focus:ring-cyan-500/20 font-mono transition-all shadow-sm" />
                </div>
              </div>
              <div className="overflow-x-auto w-full">
                <table className="w-full text-left border-collapse whitespace-nowrap text-xs">
                  <thead>
                    <tr className="bg-slate-50 dark:bg-slate-950 border-b border-slate-200 dark:border-slate-800 text-[10px] text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">
                      <th className="p-4 px-6">Date</th>
                      <th className="p-4 px-6">In</th>
                      <th className="p-4 px-6">Out</th>
                      <th className="p-4 px-6">Hours</th>
                      <th className="p-4 px-6">Remarks</th>
                      <th className="p-4 px-6 text-right font-mono">Actions</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-800/60">
                    {filteredDtr.length === 0 ? (
                      <tr><td colSpan="6" className="p-16 text-center text-slate-400 font-mono">No records found.</td></tr>
                    ) : (
                      filteredDtr.map((entry) => (
                        <tr key={entry.id} className="hover:bg-slate-50 dark:hover:bg-slate-950/40 transition-colors group">
                          <td className="p-4 px-6 font-bold text-slate-700 dark:text-slate-300 font-mono">{formatDateString(entry.date)}</td>
                          <td className="p-4 px-6 text-slate-500 dark:text-slate-400 font-mono">{entry.time_in.substring(0, 5)}</td>
                          <td className="p-4 px-6 text-slate-500 dark:text-slate-400 font-mono">{entry.time_out.substring(0, 5)}</td>
                          <td className="p-4 px-6 font-mono"><span className="px-2.5 py-1 rounded-md text-xs font-bold bg-cyan-500/10 text-cyan-600 dark:text-cyan-400">{entry.hours_rendered}h</span></td>
                          <td className="p-4 px-6 text-slate-600 dark:text-slate-300 max-w-[200px] truncate">{entry.remarks || '-'}</td>
                          <td className="p-4 px-6 text-right font-mono">
                            <div className="flex justify-end gap-2 opacity-100 sm:opacity-0 sm:group-hover:opacity-100 transition-opacity">
                              <button onClick={() => handleEditClick(entry)} className="text-cyan-600 p-2 rounded-lg bg-slate-100 dark:bg-slate-950 hover:bg-slate-200 hover:dark:bg-slate-800"><Edit className="w-4 h-4"/></button>
                              <button onClick={() => deleteEntry(entry.id)} className="text-rose-600 p-2 rounded-lg bg-slate-100 dark:bg-slate-950 hover:bg-rose-50 hover:dark:bg-rose-950/30"><Trash2 className="w-4 h-4"/></button>
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ESTIMATOR TAB */}
          {activeTab === 'estimator' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-soft max-w-2xl mx-auto w-full animate-fadeIn">
              <h3 className="text-lg font-black text-slate-900 dark:text-white mb-6 font-mono flex items-center gap-2 uppercase tracking-wide"><Calculator className="w-5 h-5 text-cyan-500"/> Completion Estimator</h3>
              <div className="space-y-6 font-mono text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                    <label className="block uppercase font-bold text-slate-400 mb-1.5 text-[10px]">Target Days / Week</label>
                    <input type="number" min="1" max="7" value={targetDaysPerWeek} onChange={(e) => setTargetDaysPerWeek(Number(e.target.value))} className="w-full bg-transparent border-none p-0 outline-none font-bold text-slate-900 dark:text-white text-lg" />
                  </div>
                  <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                    <label className="block uppercase font-bold text-slate-400 mb-1.5 text-[10px]">Target Hours / Day</label>
                    <input type="number" min="1" max="12" value={targetHoursPerDay} onChange={(e) => setTargetHoursPerDay(Number(e.target.value))} className="w-full bg-transparent border-none p-0 outline-none font-bold text-slate-900 dark:text-white text-lg" />
                  </div>
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-6 rounded-2xl border border-cyan-500/30 text-center space-y-2 shadow-inner">
                  <p className="uppercase font-bold text-slate-400 tracking-widest text-[10px]">Projected Finish Date:</p>
                  <p className="text-3xl font-black text-cyan-600 dark:text-cyan-400">
                    {remainingHours <= 0 ? 'Goal Completed! 🎉' : estimatedCompletionDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' })}
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* PROFILE TAB */}
          {activeTab === 'profile' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-soft max-w-2xl mx-auto w-full animate-fadeIn">
              <h3 className="text-lg font-black text-slate-900 dark:text-white mb-6 font-mono flex items-center gap-2 uppercase tracking-wide"><UserCheck className="w-5 h-5 text-cyan-500"/> Profile Settings</h3>
              <div className="space-y-4 font-mono text-xs">
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Full Name</label>
                  <input type="text" defaultValue={profile.full_name || ''} placeholder="Juan Dela Cruz" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent border-none p-0 outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Company / Site</label>
                  <input type="text" defaultValue={profile.company_name} onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent border-none p-0 outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">School / University</label>
                  <input type="text" defaultValue={profile.school} onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent border-none p-0 outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Required Hours Goal</label>
                  <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-full bg-transparent border-none p-0 outline-none font-bold text-cyan-600 text-sm" />
                </div>
              </div>
            </div>
          )}

        </div>
      </main>
    </div>
  );
}