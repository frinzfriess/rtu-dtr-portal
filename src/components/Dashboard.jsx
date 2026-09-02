import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import { 
  LogOut, Clock, CalendarDays, User, 
  Trash2, Edit, X, Save, FileSpreadsheet, Search, Calculator, 
  CheckCircle2, Terminal, UtensilsCrossed, LayoutDashboard, 
  FileText, UserCheck, ChevronRight, PlusCircle, Shield, Award, Sun, Moon,
  Activity, CheckCircle, Timer
} from 'lucide-react';
import { calculateHoursRendered, getGreeting, formatDateString, getCurrentTime } from '../utils';

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

  return (
    <div className="min-h-screen flex flex-col md:flex-row bg-slate-100 dark:bg-slate-950 text-slate-900 dark:text-slate-100 font-sans transition-colors duration-300 pb-28 md:pb-0 overflow-x-hidden">
      
      {editingEntry && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4 z-50">
          <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 w-full max-w-md shadow-2xl relative font-mono">
            <button onClick={() => setEditingEntry(null)} className="absolute top-6 right-6 text-slate-400 hover:text-slate-900 dark:hover:text-white bg-slate-100 dark:bg-slate-800 p-2 rounded-full">
              <X className="w-5 h-5" />
            </button>
            <h3 className="text-xl font-black text-slate-900 dark:text-white mb-6 flex items-center gap-2"><Edit className="w-5 h-5 text-cyan-500"/> Edit Record</h3>
            
            <div className="space-y-4 text-xs">
              <div>
                <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1">Date</label>
                <input type="date" value={editForm.date} onChange={e => setEditForm({...editForm, date: e.target.value})} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none font-bold text-slate-900 dark:text-white" />
              </div>
              <div className="flex gap-4">
                <div className="flex-1">
                  <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1">Time In</label>
                  <input type="time" value={editForm.time_in} onChange={e => setEditForm({...editForm, time_in: e.target.value})} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none font-bold text-slate-900 dark:text-white" />
                </div>
                <div className="flex-1">
                  <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1">Time Out</label>
                  <input type="time" value={editForm.time_out} onChange={e => setEditForm({...editForm, time_out: e.target.value})} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none font-bold text-slate-900 dark:text-white" />
                </div>
              </div>
              <div>
                <label className="block font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1">Remarks</label>
                <input type="text" value={editForm.remarks} onChange={e => setEditForm({...editForm, remarks: e.target.value})} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl outline-none font-bold text-slate-900 dark:text-white" />
              </div>
            </div>

            <div className="flex gap-3 mt-8 font-mono text-xs">
              <button onClick={() => setEditingEntry(null)} className="flex-1 bg-slate-200 dark:bg-slate-800 text-slate-800 dark:text-slate-300 font-bold py-3 rounded-xl">Cancel</button>
              <button onClick={saveEdit} disabled={isSavingEdit} className="flex-1 bg-cyan-500 text-slate-950 font-black py-3 rounded-xl flex items-center justify-center gap-2">
                <Save className="w-4 h-4"/> {isSavingEdit ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DESKTOP SIDEBAR */}
      <aside className="hidden md:flex w-72 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 flex-col justify-between p-6 flex-shrink-0">
        <div>
          <div className="flex items-center gap-3 pb-6 mb-6 border-b border-slate-200 dark:border-slate-800">
            <div className="bg-cyan-500/10 p-3 rounded-2xl text-cyan-600 dark:text-cyan-400"><Shield className="w-6 h-6" /></div>
            <div>
              <span className="text-[10px] font-mono font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest block">RTU Portal</span>
              <h1 className="text-base font-black text-slate-900 dark:text-white font-mono">DTR SYSTEM</h1>
            </div>
          </div>

          <nav className="space-y-2 font-mono text-xs">
            <button onClick={() => setActiveTab('overview')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl ${activeTab === 'overview' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><LayoutDashboard className="w-4 h-4"/> Overview</div>
              <ChevronRight className="w-3.5 h-3.5"/>
            </button>
            <button onClick={() => setActiveTab('log')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl ${activeTab === 'log' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><PlusCircle className="w-4 h-4"/> Log Hours</div>
              <ChevronRight className="w-3.5 h-3.5"/>
            </button>
            <button onClick={() => setActiveTab('history')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl ${activeTab === 'history' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><FileText className="w-4 h-4"/> History</div>
              <span className="bg-slate-200 dark:bg-slate-800 px-2 py-0.5 rounded-md font-bold">{dtr.length}</span>
            </button>
            <button onClick={() => setActiveTab('estimator')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl ${activeTab === 'estimator' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><Calculator className="w-4 h-4"/> Estimator</div>
              <ChevronRight className="w-3.5 h-3.5"/>
            </button>
            <button onClick={() => setActiveTab('profile')} className={`w-full flex items-center justify-between px-4 py-3.5 rounded-2xl ${activeTab === 'profile' ? 'bg-cyan-500 text-slate-950 font-black shadow-glow' : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
              <div className="flex items-center gap-3"><UserCheck className="w-4 h-4"/> Profile</div>
              <ChevronRight className="w-3.5 h-3.5"/>
            </button>
          </nav>
        </div>

        <div className="pt-6 border-t border-slate-200 dark:border-slate-800 space-y-4">
          <div className="flex gap-2 font-mono">
            <button onClick={exportStyledExcel} className="flex-1 flex items-center justify-center gap-1.5 bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/30 py-2.5 rounded-xl text-xs font-bold">
              <FileSpreadsheet className="w-3.5 h-3.5"/> Export
            </button>
            <button onClick={() => setDarkMode(!darkMode)} className="p-2.5 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300">
              {darkMode ? <Sun className="w-4 h-4 text-amber-400"/> : <Moon className="w-4 h-4"/>}
            </button>
            <button onClick={() => supabase.auth.signOut()} className="p-2.5 rounded-xl bg-rose-500/10 text-rose-600 border border-rose-500/30">
              <LogOut className="w-3.5 h-3.5"/>
            </button>
          </div>
        </div>
      </aside>

      {/* MOBILE BOTTOM NAV */}
      <nav aria-label="Mobile Navigation" className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-white/95 dark:bg-slate-900/95 backdrop-blur-md border-t border-slate-200 dark:border-slate-800 px-3 py-2.5 flex justify-around items-center shadow-lg">
        <button onClick={() => setActiveTab('overview')} className={`flex flex-col items-center p-1.5 ${activeTab === 'overview' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <LayoutDashboard className="w-5 h-5"/><span className="text-[10px] font-mono">Overview</span>
        </button>
        <button onClick={() => setActiveTab('log')} className={`flex flex-col items-center p-1.5 ${activeTab === 'log' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <PlusCircle className="w-5 h-5"/><span className="text-[10px] font-mono">Log</span>
        </button>
        <button onClick={() => setActiveTab('history')} className={`flex flex-col items-center p-1.5 ${activeTab === 'history' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <FileText className="w-5 h-5"/><span className="text-[10px] font-mono">History</span>
        </button>
        <button onClick={() => setActiveTab('estimator')} className={`flex flex-col items-center p-1.5 ${activeTab === 'estimator' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <Calculator className="w-5 h-5"/><span className="text-[10px] font-mono">Estimator</span>
        </button>
        <button onClick={() => setActiveTab('profile')} className={`flex flex-col items-center p-1.5 ${activeTab === 'profile' ? 'text-cyan-500 font-black' : 'text-slate-400'}`}>
          <UserCheck className="w-5 h-5"/><span className="text-[10px] font-mono">Profile</span>
        </button>
      </nav>

      {/* MAIN CONTAINER */}
      <main className="flex-1 flex flex-col min-w-0 w-full overflow-y-auto bg-slate-100 dark:bg-slate-950">
        
        <header className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-md border-b border-slate-200 dark:border-slate-800 px-4 sm:px-6 py-4 flex items-center justify-between sticky top-0 z-30">
          <div className="min-w-0 pr-2">
            <div className="flex items-center gap-2 text-cyan-600 dark:text-cyan-400 text-[11px] font-mono font-bold uppercase tracking-widest">
              <span className="w-2 h-2 rounded-full bg-cyan-500 animate-pulse"></span>
              <span>{getGreeting()}</span>
            </div>
            <h2 className="text-base sm:text-xl font-black text-slate-900 dark:text-white truncate">
              {activeTab === 'overview' && 'Dashboard Overview'}
              {activeTab === 'log' && 'Log Daily Hours'}
              {activeTab === 'history' && 'Timesheet History'}
              {activeTab === 'estimator' && 'OJT Estimator'}
              {activeTab === 'profile' && 'Student Profile'}
            </h2>
          </div>

          <div className="flex items-center gap-2 font-mono text-xs">
            <button onClick={() => setDarkMode(!darkMode)} className="md:hidden p-2 rounded-xl bg-slate-200 dark:bg-slate-800 text-slate-800 dark:text-slate-200">
              {darkMode ? <Sun className="w-4 h-4 text-amber-400"/> : <Moon className="w-4 h-4"/>}
            </button>
            <button onClick={exportStyledExcel} className="hidden sm:flex items-center gap-1.5 bg-emerald-600 text-white font-bold px-3 py-2 rounded-xl">
              <FileSpreadsheet className="w-4 h-4"/> Export
            </button>
            <button onClick={() => supabase.auth.signOut()} className="hidden sm:flex items-center gap-1.5 bg-white dark:bg-slate-900 text-rose-600 border border-slate-200 dark:border-slate-800 font-bold px-3 py-2 rounded-xl">
              <LogOut className="w-4 h-4"/>
            </button>
          </div>
        </header>

        <div className="p-3 sm:p-6 lg:p-8 max-w-5xl w-full mx-auto space-y-6">

          {/* OVERVIEW TAB */}
          {activeTab === 'overview' && (
            <div className="space-y-6 w-full">
              <div className="bg-gradient-to-r from-cyan-600 via-teal-600 to-indigo-600 text-white rounded-3xl p-6 sm:p-8 shadow-xl relative overflow-hidden flex flex-col justify-between">
                <div className="relative z-10">
                  <span className="bg-white/20 px-3 py-1 rounded-full text-[10px] font-mono font-bold uppercase backdrop-blur-md">RTU OJT Portal</span>
                  <h3 className="text-2xl sm:text-3xl font-black tracking-tight mt-3 mb-2">Welcome, {profile.full_name || 'Student'}!</h3>
                  <p className="text-xs sm:text-sm text-cyan-100 max-w-xl mb-6">
                    Completed <strong className="text-white">{progressPercentage.toFixed(1)}%</strong> of your required hours at <strong className="text-white">{profile.company_name}</strong>.
                  </p>
                </div>
                <div className="relative z-10 flex flex-wrap items-center gap-3 font-mono">
                  <button onClick={() => setActiveTab('log')} className="bg-white text-slate-950 font-black px-5 py-3 rounded-2xl text-xs flex items-center gap-2">
                    <PlusCircle className="w-4 h-4 text-cyan-600"/> Log Hours
                  </button>
                  <button onClick={exportStyledExcel} className="bg-black/20 text-white font-bold px-5 py-3 rounded-2xl text-xs border border-white/20">
                    Export DTR
                  </button>
                </div>
              </div>

              {/* LIVE TIMER */}
              <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft flex flex-col sm:flex-row items-center justify-between gap-4 w-full">
                <div className="flex items-center gap-3.5">
                  <div className="bg-cyan-500/10 text-cyan-600 dark:text-cyan-400 p-3 rounded-2xl border border-cyan-500/20"><Timer className="w-6 h-6 animate-pulse"/></div>
                  <div>
                    <h4 className="text-xs font-mono font-bold text-slate-400 uppercase tracking-widest">Shift Stopwatch</h4>
                    <p className="text-2xl font-black font-mono text-slate-900 dark:text-white">
                      {Math.floor(timerSeconds / 3600).toString().padStart(2, '0')}:
                      {Math.floor((timerSeconds % 3600) / 60).toString().padStart(2, '0')}:
                      {(timerSeconds % 60).toString().padStart(2, '0')}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-2 font-mono text-xs w-full sm:w-auto">
                  {!isTimerRunning ? (
                    <button onClick={() => setIsTimerRunning(true)} className="flex-1 sm:flex-none bg-emerald-500 text-slate-950 font-black px-5 py-3 rounded-xl">Start</button>
                  ) : (
                    <button onClick={() => setIsTimerRunning(false)} className="flex-1 sm:flex-none bg-rose-500 text-white font-black px-5 py-3 rounded-xl">Pause</button>
                  )}
                  <button onClick={() => { setIsTimerRunning(false); setTimerSeconds(0); }} className="bg-slate-200 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold px-4 py-3 rounded-xl">Reset</button>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 w-full">
                <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft">
                  <span className="text-[10px] font-mono font-bold text-slate-400 uppercase">Rendered</span>
                  <h3 className="text-3xl font-black font-mono mt-1 mb-2">{totalHours.toFixed(2)}<span className="text-sm font-bold text-cyan-500"> hrs</span></h3>
                  <div className="w-full bg-slate-100 dark:bg-slate-950 rounded-full h-2 overflow-hidden border border-slate-200 dark:border-slate-800">
                    <div className="bg-cyan-500 h-full rounded-full" style={{ width: `${progressPercentage}%` }}></div>
                  </div>
                </div>
                <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft">
                  <span className="text-[10px] font-mono font-bold text-slate-400 uppercase">Shifts Logged</span>
                  <h3 className="text-3xl font-black font-mono mt-1 mb-2">{dtr.length}</h3>
                  <span className="text-xs font-mono text-cyan-500 font-bold">100% Attendance</span>
                </div>
                <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 p-5 rounded-3xl shadow-soft">
                  <span className="text-[10px] font-mono font-bold text-slate-400 uppercase">Remaining</span>
                  <h3 className="text-3xl font-black font-mono mt-1 mb-2 text-emerald-500">{remainingHours.toFixed(2)}<span className="text-sm font-bold"> hrs</span></h3>
                  <span className="text-xs font-mono text-emerald-500 font-bold">{remainingHours <= 0 ? 'Completed 🎉' : 'On Track'}</span>
                </div>
              </div>
            </div>
          )}

          {/* ABSOLUTELY FIXED NON-OVERLAPPING LOG HOURS TAB */}
          {activeTab === 'log' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-5 sm:p-8 shadow-soft max-w-xl mx-auto w-full box-border overflow-hidden">
              
              <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-100 dark:border-slate-800 w-full">
                <div className="bg-cyan-500/10 p-3 rounded-2xl text-cyan-600 dark:text-cyan-400 border border-cyan-500/30 flex-shrink-0">
                  <Terminal className="w-6 h-6" />
                </div>
                <div className="min-w-0 flex-1">
                  <h3 className="text-base sm:text-lg font-black text-slate-900 dark:text-white uppercase tracking-wider font-mono truncate">Log Daily Hours</h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400 font-mono truncate">Record your shift times safely.</p>
                </div>
              </div>

              <form onSubmit={handleLogSubmit} className="space-y-4 w-full block">
                
                {/* LUNCH BREAK TOGGLE */}
                <div className="w-full">
                  <button
                    type="button"
                    onClick={() => setDeductLunch(!deductLunch)}
                    className={`w-full flex items-center justify-between px-4 py-3 rounded-2xl border transition-all select-none ${
                      deductLunch ? 'bg-cyan-500/10 border-cyan-500/40 text-cyan-600 dark:text-cyan-400' : 'bg-slate-50 dark:bg-slate-950 border-slate-200 dark:border-slate-800 text-slate-400'
                    }`}
                  >
                    <div className="flex items-center gap-2 min-w-0">
                      <UtensilsCrossed className="w-4 h-4 flex-shrink-0 text-cyan-500" />
                      <span className="text-xs font-bold font-mono truncate">Auto-deduct 1h lunch break</span>
                    </div>
                    <div className={`w-5 h-5 rounded-full flex items-center justify-center border flex-shrink-0 ${
                      deductLunch ? 'bg-cyan-500 border-cyan-600 text-slate-950 font-bold' : 'bg-transparent border-slate-400'
                    }`}>
                      {deductLunch && <div className="w-2 h-2 rounded-full bg-slate-950"></div>}
                    </div>
                  </button>
                </div>

                {/* DATE CONTAINER */}
                <div className="w-full">
                  <label className="block text-[11px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Date</label>
                  <input type="date" required value={date} onChange={(e) => setDate(e.target.value)} 
                    className="w-full px-4 py-3.5 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-2xl text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 font-mono block box-border shadow-sm" />
                </div>

                {/* TIME IN CONTAINER */}
                <div className="w-full bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <div className="flex justify-between items-center mb-2 w-full">
                     <label className="text-[11px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time In</label>
                     <button type="button" onClick={() => setTimeIn(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2.5 py-1 rounded-lg font-mono flex items-center gap-1">
                       <Clock className="w-3 h-3"/> Now
                     </button>
                  </div>
                  <input type="time" required value={timeIn} onChange={(e) => setTimeIn(e.target.value)} 
                    className="w-full px-3 py-3 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl text-base font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 font-mono text-center block box-border" />
                </div>

                {/* TIME OUT CONTAINER */}
                <div className="w-full bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <div className="flex justify-between items-center mb-2 w-full">
                     <label className="text-[11px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest font-mono">Time Out</label>
                     <button type="button" onClick={() => setTimeOut(getCurrentTime())} className="text-[10px] font-bold text-cyan-600 dark:text-cyan-400 bg-cyan-500/10 px-2.5 py-1 rounded-lg font-mono flex items-center gap-1">
                       <Clock className="w-3 h-3"/> Now
                     </button>
                  </div>
                  <input type="time" required value={timeOut} onChange={(e) => setTimeOut(e.target.value)} 
                    className="w-full px-3 py-3 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl text-base font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 font-mono text-center block box-border" />
                </div>

                {/* REMARKS CONTAINER */}
                <div className="w-full">
                  <label className="block text-[11px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1.5 font-mono">Remarks / Tasks</label>
                  <textarea rows="3" placeholder="E.g., System debugging, UI updates..." value={remarks} onChange={(e) => setRemarks(e.target.value)} 
                    className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-2xl text-xs sm:text-sm font-bold text-slate-900 dark:text-white outline-none focus:border-cyan-500 font-sans resize-none block box-border shadow-sm" />
                </div>

                {/* SUBMIT BUTTON */}
                <div className="pt-2 w-full">
                  <button type="submit" disabled={submitting} 
                    className={`w-full font-bold py-4 rounded-2xl transition-all shadow-glow font-mono text-xs tracking-wide flex items-center justify-center gap-2 block box-border ${
                      successStatus ? 'bg-emerald-500 text-white shadow-neon' : 'bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black'
                    }`}>
                    {successStatus ? <><CheckCircle2 className="w-4 h-4"/> Logged Successfully!</> : 'Submit Daily Record'}
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* HISTORY TAB */}
          {activeTab === 'history' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl shadow-soft overflow-hidden w-full">
              <div className="px-4 sm:px-6 py-5 border-b border-slate-100 dark:border-slate-800 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-slate-50 dark:bg-slate-950/50">
                <h3 className="font-black text-slate-900 dark:text-white flex items-center gap-2 text-base font-mono uppercase"><Clock className="w-5 h-5 text-cyan-500"/> History</h3>
                <div className="relative w-full sm:w-72">
                   <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none"><Search className="h-4 w-4 text-slate-400" /></div>
                   <input type="text" placeholder="Search..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} className="w-full pl-10 pr-4 py-2.5 bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl text-xs font-bold text-slate-900 dark:text-white outline-none" />
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
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                    {filteredDtr.length === 0 ? (
                      <tr><td colSpan="6" className="p-12 text-center text-slate-400 font-mono">No records found.</td></tr>
                    ) : (
                      filteredDtr.map((entry) => (
                        <tr key={entry.id} className="hover:bg-slate-50 dark:hover:bg-slate-950/40 transition-colors group">
                          <td className="p-4 px-6 font-bold text-slate-700 dark:text-slate-300 font-mono">{formatDateString(entry.date)}</td>
                          <td className="p-4 px-6 text-slate-500 dark:text-slate-400 font-mono">{entry.time_in.substring(0, 5)}</td>
                          <td className="p-4 px-6 text-slate-500 dark:text-slate-400 font-mono">{entry.time_out.substring(0, 5)}</td>
                          <td className="p-4 px-6 font-mono"><span className="px-2.5 py-1 rounded-md text-xs font-bold bg-cyan-500/10 text-cyan-600">{entry.hours_rendered}h</span></td>
                          <td className="p-4 px-6 text-slate-600 dark:text-slate-300 max-w-[200px] truncate">{entry.remarks || '-'}</td>
                          <td className="p-4 px-6 text-right font-mono">
                            <div className="flex justify-end gap-2">
                              <button onClick={() => handleEditClick(entry)} className="text-cyan-600 p-2 rounded-lg bg-slate-100 dark:bg-slate-950"><Edit className="w-4 h-4"/></button>
                              <button onClick={() => deleteEntry(entry.id)} className="text-rose-600 p-2 rounded-lg bg-slate-100 dark:bg-slate-950"><Trash2 className="w-4 h-4"/></button>
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
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-soft max-w-2xl mx-auto w-full">
              <h3 className="text-lg font-black text-slate-900 dark:text-white mb-4 font-mono flex items-center gap-2"><Calculator className="w-5 h-5 text-cyan-500"/> Completion Estimator</h3>
              <div className="space-y-5 font-mono text-xs">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block uppercase font-bold text-slate-400 mb-1">Days / Week</label>
                    <input type="number" min="1" max="7" value={targetDaysPerWeek} onChange={(e) => setTargetDaysPerWeek(Number(e.target.value))} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-bold text-slate-900 dark:text-white" />
                  </div>
                  <div>
                    <label className="block uppercase font-bold text-slate-400 mb-1">Hours / Day</label>
                    <input type="number" min="1" max="12" value={targetHoursPerDay} onChange={(e) => setTargetHoursPerDay(Number(e.target.value))} className="w-full px-4 py-3 bg-slate-50 dark:bg-slate-950 border border-slate-200 dark:border-slate-800 rounded-xl font-bold text-slate-900 dark:text-white" />
                  </div>
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-6 rounded-2xl border border-cyan-500/30 text-center space-y-2">
                  <p className="uppercase font-bold text-slate-400">Projected Finish Date:</p>
                  <p className="text-2xl font-black text-cyan-600 dark:text-cyan-400">
                    {remainingHours <= 0 ? 'Goal Completed! 🎉' : estimatedCompletionDate.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' })}
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* PROFILE TAB */}
          {activeTab === 'profile' && (
            <div className="bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-6 sm:p-8 shadow-soft max-w-2xl mx-auto w-full">
              <h3 className="text-lg font-black text-slate-900 dark:text-white mb-6 font-mono flex items-center gap-2"><User className="w-5 h-5 text-cyan-500"/> Profile Settings</h3>
              <div className="space-y-4 font-mono text-xs">
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Full Name</label>
                  <input type="text" defaultValue={profile.full_name || ''} placeholder="Juan Dela Cruz" onBlur={(e) => updateProfileInfo('full_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Company Site</label>
                  <input type="text" defaultValue={profile.company_name} onBlur={(e) => updateProfileInfo('company_name', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">School</label>
                  <input type="text" defaultValue={profile.school} onBlur={(e) => updateProfileInfo('school', e.target.value)} className="w-full bg-transparent outline-none font-bold text-slate-900 dark:text-white text-sm" />
                </div>
                <div className="bg-slate-50 dark:bg-slate-950 p-4 rounded-2xl border border-slate-200 dark:border-slate-800">
                  <label className="text-[10px] uppercase text-slate-400 font-bold block mb-1">Required Hours Goal</label>
                  <input type="number" defaultValue={profile.required_hours} onBlur={(e) => updateProfileInfo('required_hours', e.target.value)} className="w-full bg-transparent outline-none font-bold text-cyan-600 text-sm" />
                </div>
              </div>
            </div>
          )}

        </div>
      </main>
    </div>
  );
}
