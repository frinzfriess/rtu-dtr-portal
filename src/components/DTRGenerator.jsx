import { useEffect, useMemo, useState } from 'react';
import { Printer, RotateCcw } from 'lucide-react';
import { getHoursDiff, getDaysInMonth, formatMonthLabel } from '../utils';

function emptyRow() {
  return { am_in: '', am_out: '', pm_in: '', pm_out: '' };
}

function buildRowsFromLogs(dtr, month, dayCount) {
  const rows = Array.from({ length: dayCount }, emptyRow);
  (dtr || []).forEach((entry) => {
    if (!entry.date || !entry.date.startsWith(month)) return;
    const day = Number(entry.date.split('-')[2]);
    if (!day || day < 1 || day > dayCount) return;

    const timeIn = entry.time_in?.substring(0, 5) || '';
    const timeOut = entry.time_out?.substring(0, 5) || '';
    if (!timeIn || !timeOut) return;

    // Shifts that cross the lunch hour are split into AM/PM halves (matches the
    // 1h lunch auto-deduction already used for the logged hours elsewhere in the app).
    if (timeIn < '12:00' && timeOut > '13:00') {
      rows[day - 1] = { am_in: timeIn, am_out: '12:00', pm_in: '13:00', pm_out: timeOut };
    } else {
      rows[day - 1] = { am_in: timeIn, am_out: timeOut, pm_in: '', pm_out: '' };
    }
  });
  return rows;
}

export default function DTRGenerator({ profile, dtr }) {
  const [month, setMonth] = useState(() => new Date().toISOString().slice(0, 7));
  const [office, setOffice] = useState('OFFICE OF THE ADJUTANT, PA');
  const [name, setName] = useState(profile?.full_name || '');
  const [regularHours, setRegularHours] = useState('8:00 AM - 5:00 PM');
  const [saturdayHours, setSaturdayHours] = useState('8:00 AM - 12:00 PM');
  const [verifierName, setVerifierName] = useState('ARMI L. AGATEA, MBA');
  const [verifierTitle, setVerifierTitle] = useState('Admin Officer V');
  const [verifierOffice, setVerifierOffice] = useState('Civ HR Supervisor');

  const dayCount = useMemo(() => getDaysInMonth(month), [month]);
  const [rows, setRows] = useState(() => buildRowsFromLogs(dtr, month, dayCount));

  useEffect(() => {
    setRows(buildRowsFromLogs(dtr, month, dayCount));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [month, dayCount]);

  function refreshFromLogs() {
    setRows(buildRowsFromLogs(dtr, month, dayCount));
  }

  function updateCell(dayIdx, field, value) {
    setRows((prev) => prev.map((r, i) => (i === dayIdx ? { ...r, [field]: value } : r)));
  }

  const dailyTotals = useMemo(
    () => rows.map((r) => getHoursDiff(r.am_in, r.am_out) + getHoursDiff(r.pm_in, r.pm_out)),
    [rows]
  );
  const monthTotal = useMemo(() => dailyTotals.reduce((sum, h) => sum + h, 0), [dailyTotals]);
  const monthLabel = formatMonthLabel(month);

  return (
    <div className="w-full animate-fadeIn">
      {/* CONTROLS (hidden when printing) */}
      <div className="print:hidden bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-4 sm:p-5 shadow-soft mb-6 flex flex-wrap items-end gap-4 font-mono">
        <div className="bg-slate-50 dark:bg-slate-950 px-4 py-2.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
          <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 font-mono">Month</label>
          <input
            type="month"
            value={month}
            onChange={(e) => setMonth(e.target.value)}
            className="w-full bg-transparent border-none p-0 m-0 outline-none text-sm font-bold text-slate-900 dark:text-white block appearance-none"
          />
        </div>
        <button
          onClick={refreshFromLogs}
          className="flex items-center gap-1.5 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 font-bold px-4 py-3 rounded-xl text-xs transition-all active:scale-[0.96]"
        >
          <RotateCcw className="w-3.5 h-3.5" /> Refresh from Logged Hours
        </button>
        <button
          onClick={() => window.print()}
          className="flex items-center gap-1.5 bg-cyan-500 hover:bg-cyan-400 text-slate-950 font-black px-5 py-3 rounded-xl text-xs shadow-glow transition-all active:scale-[0.96] ml-auto"
        >
          <Printer className="w-4 h-4" /> Print / Save as PDF
        </button>
        <p className="text-[10px] text-slate-400 w-full font-mono leading-relaxed">
          Fill in Name, Office and the verifying officer's details below — they're editable directly on the sheet. Time cells are pre-filled from your Logged Hours for the selected month and can be edited by hand.
        </p>
      </div>

      {/* PRINTABLE SHEET */}
      <div className="dtr-sheet bg-white text-slate-900 border border-slate-200 rounded-3xl shadow-soft print:shadow-none print:border-0 print:rounded-none p-6 sm:p-10 max-w-[850px] mx-auto font-serif print:p-0">
        <div className="text-center mb-5 print:mb-2">
          <h1 className="text-xl font-bold uppercase tracking-wide print:text-lg">Daily Time Record</h1>
          <input
            type="text"
            value={office}
            onChange={(e) => setOffice(e.target.value)}
            placeholder="Office / Unit"
            className="mt-1 block mx-auto w-full max-w-md text-center text-sm font-semibold uppercase bg-transparent border-0 outline-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent"
          />
        </div>

        <div className="flex flex-wrap justify-between items-end gap-4 mb-3 print:mb-1 text-sm">
          <div className="flex-1 min-w-[220px]">
            <span className="font-semibold">Name: </span>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="Juan Dela Cruz"
              className="border-b border-slate-500 outline-none bg-transparent font-bold px-1 w-[65%] transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent print:border-black"
            />
          </div>
          <div className="text-right">
            <span className="font-semibold">Month of: </span>
            <span className="font-bold border-b border-slate-500 px-1 print:border-black">{monthLabel}</span>
          </div>
        </div>

        <div className="text-xs mb-4 print:mb-2 leading-relaxed">
          <p className="font-semibold">Official hours for arrival and departure</p>
          <div className="flex flex-wrap gap-6 mt-1">
            <div>
              <span className="font-semibold">Regular Days: </span>
              <input
                type="text"
                value={regularHours}
                onChange={(e) => setRegularHours(e.target.value)}
                className="border-b border-slate-400 outline-none bg-transparent px-1 w-44 transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent print:border-black"
              />
            </div>
            <div>
              <span className="font-semibold">Saturdays: </span>
              <input
                type="text"
                value={saturdayHours}
                onChange={(e) => setSaturdayHours(e.target.value)}
                className="border-b border-slate-400 outline-none bg-transparent px-1 w-44 transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent print:border-black"
              />
            </div>
          </div>
        </div>

        <table className="w-full border-collapse text-[11px] print:text-[9px] mb-1">
          <thead>
            <tr>
              <th rowSpan={2} className="border border-slate-600 px-1 py-1 print:py-0.5 w-8 print:border-black">Day</th>
              <th colSpan={2} className="border border-slate-600 px-1 py-1 print:py-0.5 print:border-black">A.M.</th>
              <th colSpan={2} className="border border-slate-600 px-1 py-1 print:py-0.5 print:border-black">P.M.</th>
              <th rowSpan={2} className="border border-slate-600 px-1 py-1 print:py-0.5 w-16 print:border-black">Total Hours</th>
            </tr>
            <tr>
              <th className="border border-slate-600 px-1 py-0.5 print:border-black">Arrival</th>
              <th className="border border-slate-600 px-1 py-0.5 print:border-black">Departure</th>
              <th className="border border-slate-600 px-1 py-0.5 print:border-black">Arrival</th>
              <th className="border border-slate-600 px-1 py-0.5 print:border-black">Departure</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx}>
                <td className="border border-slate-400 text-center font-semibold print:border-black">{idx + 1}</td>
                <td className="border border-slate-400 p-0 print:border-black">
                  <input type="time" value={row.am_in} onChange={(e) => updateCell(idx, 'am_in', e.target.value)}
                    className="w-full text-center bg-transparent outline-none py-0.5 print:py-0 text-[10px] print:text-[9px] print:h-[15px] print:leading-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent" />
                </td>
                <td className="border border-slate-400 p-0 print:border-black">
                  <input type="time" value={row.am_out} onChange={(e) => updateCell(idx, 'am_out', e.target.value)}
                    className="w-full text-center bg-transparent outline-none py-0.5 print:py-0 text-[10px] print:text-[9px] print:h-[15px] print:leading-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent" />
                </td>
                <td className="border border-slate-400 p-0 print:border-black">
                  <input type="time" value={row.pm_in} onChange={(e) => updateCell(idx, 'pm_in', e.target.value)}
                    className="w-full text-center bg-transparent outline-none py-0.5 print:py-0 text-[10px] print:text-[9px] print:h-[15px] print:leading-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent" />
                </td>
                <td className="border border-slate-400 p-0 print:border-black">
                  <input type="time" value={row.pm_out} onChange={(e) => updateCell(idx, 'pm_out', e.target.value)}
                    className="w-full text-center bg-transparent outline-none py-0.5 print:py-0 text-[10px] print:text-[9px] print:h-[15px] print:leading-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent" />
                </td>
                <td className="border border-slate-400 text-center font-semibold print:border-black">
                  {dailyTotals[idx] > 0 ? dailyTotals[idx].toFixed(2) : ''}
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr>
              <td colSpan={5} className="border border-slate-600 text-right font-bold px-2 py-1.5 print:py-0.5 print:border-black">
                TOTAL HOURS FOR {monthLabel ? monthLabel.toUpperCase() : 'THE MONTH'}
              </td>
              <td className="border border-slate-600 text-center font-bold py-1.5 print:py-0.5 print:border-black">{monthTotal.toFixed(2)}</td>
            </tr>
          </tfoot>
        </table>

        <p className="text-[10px] mt-4 print:mt-2 leading-relaxed">
          I CERTIFY on my honor that the above is a true and correct report of hours of work performed, record of which was made daily at the time of arrival and departure from office.
        </p>

        <div className="flex flex-wrap justify-between gap-8 mt-10 print:mt-6 text-center text-[10px]">
          <div className="flex-1 min-w-[220px]">
            <div className="border-t border-slate-600 pt-1 font-bold print:border-black">{name || ' '}</div>
            <p className="mt-0.5">Signature of Employee</p>
          </div>

          <div className="flex-1 min-w-[220px]">
            <p className="font-semibold mb-8 print:mb-4 text-left">Verified as to prescribed office hours:</p>
            <input
              type="text"
              value={verifierName}
              onChange={(e) => setVerifierName(e.target.value)}
              className="border-t border-slate-600 pt-1 font-bold text-center bg-transparent outline-none w-full transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent print:border-black"
            />
            <input
              type="text"
              value={verifierTitle}
              onChange={(e) => setVerifierTitle(e.target.value)}
              className="text-center bg-transparent outline-none w-full mt-0.5 transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent"
            />
            <input
              type="text"
              value={verifierOffice}
              onChange={(e) => setVerifierOffice(e.target.value)}
              className="text-center bg-transparent outline-none w-full mt-0.5 transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent"
            />
          </div>
        </div>
      </div>
    </div>
  );
}
