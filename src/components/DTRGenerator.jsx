import { useEffect, useMemo, useState } from 'react';
import { Printer, RotateCcw } from 'lucide-react';
import { getHoursDiff, formatMonthLabel, shiftMonth } from '../utils';

const DAYS_IN_FORM = 31;

function emptyRow() {
  return { am_in: '', am_out: '', pm_in: '', pm_out: '' };
}

function buildRowsFromLogs(dtr, month, dayCount = DAYS_IN_FORM) {
  const rows = Array.from({ length: dayCount }, emptyRow);
  (dtr || []).forEach((entry) => {
    if (!entry.date || !entry.date.startsWith(month)) return;
    const day = Number(entry.date.split('-')[2]);
    if (!day || day < 1 || day > dayCount) return;

    const timeIn = entry.time_in?.substring(0, 5) || '';
    const timeOut = entry.time_out?.substring(0, 5) || '';
    if (!timeIn || !timeOut) return;

    // No lunch break to carve out — the logged time in/out is the full worked
    // span, so it goes straight into one pair and the total matches the log exactly.
    rows[day - 1] = { am_in: timeIn, am_out: timeOut, pm_in: '', pm_out: '' };
  });
  return rows;
}

const inputCls = 'w-full bg-transparent border-0 outline-none transition-colors duration-200 focus:bg-amber-50 print:focus:bg-transparent';

function DTRSheet({
  office, setOffice, name, setName, regularHours, setRegularHours, saturdayHours, setSaturdayHours,
  verifierName, setVerifierName, verifierTitle, setVerifierTitle, verifierOffice, setVerifierOffice,
  month, setMonth, rows, updateCell, dailyTotals, monthTotal, monthLabel, idPrefix
}) {
  return (
    <div className="dtr-sheet bg-white text-slate-900 border border-slate-200 rounded-2xl shadow-soft print:shadow-none print:border-0 print:rounded-none p-4 sm:p-5 font-serif print:p-0 w-full">
      <div className="text-center mb-2">
        <h1 className="text-sm font-bold uppercase tracking-wide print:text-[11px]">Daily Time Record</h1>
        <input
          type="text"
          value={office}
          onChange={(e) => setOffice(e.target.value)}
          placeholder="Office / Unit"
          className={`mt-0.5 block mx-auto text-center text-[9px] font-semibold uppercase ${inputCls}`}
        />
      </div>

      <div className="mb-1.5 text-[9px]">
        <span className="font-semibold">Name: </span>
        <input
          type="text"
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Juan Dela Cruz"
          className={`border-b border-slate-500 font-bold px-1 w-[62%] print:border-black ${inputCls}`}
        />
      </div>

      <div className="mb-1.5 text-[9px]">
        <span className="font-semibold">Month of: </span>
        <input
          type="month"
          value={month}
          onChange={(e) => setMonth(e.target.value)}
          className={`border-b border-slate-500 font-bold px-1 print:border-black ${inputCls}`}
        />
      </div>

      <div className="text-[8px] mb-1.5 leading-tight">
        <p className="font-semibold">Official hours for arrival and departure</p>
        <div className="flex flex-wrap gap-x-3">
          <span>
            <span className="font-semibold">Reg: </span>
            <input type="text" value={regularHours} onChange={(e) => setRegularHours(e.target.value)}
              className={`border-b border-slate-400 px-0.5 w-20 print:border-black ${inputCls}`} />
          </span>
          <span>
            <span className="font-semibold">Sat: </span>
            <input type="text" value={saturdayHours} onChange={(e) => setSaturdayHours(e.target.value)}
              className={`border-b border-slate-400 px-0.5 w-20 print:border-black ${inputCls}`} />
          </span>
        </div>
      </div>

      <table className="w-full border-collapse text-[7.5px] print:text-[7px] mb-1">
        <thead>
          <tr>
            <th rowSpan={2} className="border border-slate-600 px-0.5 py-0.5 w-4 print:border-black">Day</th>
            <th colSpan={2} className="border border-slate-600 px-0.5 py-0.5 print:border-black">A.M.</th>
            <th colSpan={2} className="border border-slate-600 px-0.5 py-0.5 print:border-black">P.M.</th>
            <th rowSpan={2} className="border border-slate-600 px-0.5 py-0.5 w-7 print:border-black">Total</th>
          </tr>
          <tr>
            <th className="border border-slate-600 px-0.5 py-0.5 print:border-black">In</th>
            <th className="border border-slate-600 px-0.5 py-0.5 print:border-black">Out</th>
            <th className="border border-slate-600 px-0.5 py-0.5 print:border-black">In</th>
            <th className="border border-slate-600 px-0.5 py-0.5 print:border-black">Out</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row, idx) => (
            <tr key={`${idPrefix}-${idx}`}>
              <td className="border border-slate-400 text-center font-semibold print:border-black">{idx + 1}</td>
              <td className="border border-slate-400 p-0 print:border-black">
                <input type="time" value={row.am_in} onChange={(e) => updateCell(idx, 'am_in', e.target.value)}
                  className={`text-center py-0 print:h-[11px] print:leading-none ${inputCls}`} />
              </td>
              <td className="border border-slate-400 p-0 print:border-black">
                <input type="time" value={row.am_out} onChange={(e) => updateCell(idx, 'am_out', e.target.value)}
                  className={`text-center py-0 print:h-[11px] print:leading-none ${inputCls}`} />
              </td>
              <td className="border border-slate-400 p-0 print:border-black">
                <input type="time" value={row.pm_in} onChange={(e) => updateCell(idx, 'pm_in', e.target.value)}
                  className={`text-center py-0 print:h-[11px] print:leading-none ${inputCls}`} />
              </td>
              <td className="border border-slate-400 p-0 print:border-black">
                <input type="time" value={row.pm_out} onChange={(e) => updateCell(idx, 'pm_out', e.target.value)}
                  className={`text-center py-0 print:h-[11px] print:leading-none ${inputCls}`} />
              </td>
              <td className="border border-slate-400 text-center font-semibold print:border-black">
                {dailyTotals[idx] > 0 ? dailyTotals[idx].toFixed(2) : ''}
              </td>
            </tr>
          ))}
        </tbody>
        <tfoot>
          <tr>
            <td colSpan={5} className="border border-slate-600 text-right font-bold px-1 py-0.5 print:border-black">
              TOTAL FOR {monthLabel ? monthLabel.toUpperCase() : 'THE MONTH'}
            </td>
            <td className="border border-slate-600 text-center font-bold py-0.5 print:border-black">{monthTotal.toFixed(2)}</td>
          </tr>
        </tfoot>
      </table>

      <p className="text-[6.5px] mt-1.5 leading-snug">
        I CERTIFY on my honor that the above is a true and correct report of hours of work performed, record of which was made daily at the time of arrival and departure from office.
      </p>

      <div className="flex flex-col items-center text-center text-[7px]">
        <div className="mt-6 print:mt-5 w-40">
          <div className="border-t border-slate-600 pt-0.5 font-bold print:border-black">{name || ' '}</div>
          <p>Signature of Employee</p>
        </div>

        <div className="mt-6 print:mt-5 w-40">
          <p className="font-semibold mb-3 print:mb-2.5 text-center">Verified as to prescribed office hours:</p>
          <input
            type="text"
            value={verifierName}
            onChange={(e) => setVerifierName(e.target.value)}
            className={`border-t border-slate-600 pt-0.5 font-bold text-center print:border-black ${inputCls}`}
          />
          <input
            type="text"
            value={verifierTitle}
            onChange={(e) => setVerifierTitle(e.target.value)}
            className={`text-center mt-0.5 ${inputCls}`}
          />
          <input
            type="text"
            value={verifierOffice}
            onChange={(e) => setVerifierOffice(e.target.value)}
            className={`text-center mt-0.5 ${inputCls}`}
          />
        </div>
      </div>
    </div>
  );
}

export default function DTRGenerator({ profile, dtr }) {
  const currentMonth = new Date().toISOString().slice(0, 7);
  const [monthA, setMonthA] = useState(currentMonth);
  const [monthB, setMonthB] = useState(() => shiftMonth(currentMonth, -1));

  const [office, setOffice] = useState('OFFICE OF THE ADJUTANT, PA');
  const [name, setName] = useState(profile?.full_name || '');
  const [regularHours, setRegularHours] = useState('8:00 AM - 5:00 PM');
  const [saturdayHours, setSaturdayHours] = useState('8:00 AM - 12:00 PM');
  const [verifierName, setVerifierName] = useState('ARMI L. AGATEA, MBA');
  const [verifierTitle, setVerifierTitle] = useState('Admin Officer V');
  const [verifierOffice, setVerifierOffice] = useState('Civ HR Supervisor');

  const [rowsA, setRowsA] = useState(() => buildRowsFromLogs(dtr, monthA));
  const [rowsB, setRowsB] = useState(() => buildRowsFromLogs(dtr, monthB));

  useEffect(() => {
    setRowsA(buildRowsFromLogs(dtr, monthA));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [monthA]);

  useEffect(() => {
    setRowsB(buildRowsFromLogs(dtr, monthB));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [monthB]);

  function refreshFromLogs() {
    setRowsA(buildRowsFromLogs(dtr, monthA));
    setRowsB(buildRowsFromLogs(dtr, monthB));
  }

  function updateCellA(dayIdx, field, value) {
    setRowsA((prev) => prev.map((r, i) => (i === dayIdx ? { ...r, [field]: value } : r)));
  }
  function updateCellB(dayIdx, field, value) {
    setRowsB((prev) => prev.map((r, i) => (i === dayIdx ? { ...r, [field]: value } : r)));
  }

  const dailyTotalsA = useMemo(() => rowsA.map((r) => getHoursDiff(r.am_in, r.am_out) + getHoursDiff(r.pm_in, r.pm_out)), [rowsA]);
  const dailyTotalsB = useMemo(() => rowsB.map((r) => getHoursDiff(r.am_in, r.am_out) + getHoursDiff(r.pm_in, r.pm_out)), [rowsB]);
  const monthTotalA = useMemo(() => dailyTotalsA.reduce((sum, h) => sum + h, 0), [dailyTotalsA]);
  const monthTotalB = useMemo(() => dailyTotalsB.reduce((sum, h) => sum + h, 0), [dailyTotalsB]);

  const sharedFieldProps = {
    office, setOffice, name, setName, regularHours, setRegularHours, saturdayHours, setSaturdayHours,
    verifierName, setVerifierName, verifierTitle, setVerifierTitle, verifierOffice, setVerifierOffice
  };

  return (
    <div className="w-full animate-fadeIn">
      {/* CONTROLS (hidden when printing) */}
      <div className="print:hidden bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl p-4 sm:p-5 shadow-soft mb-6 flex flex-wrap items-end gap-4 font-mono">
        <div className="bg-slate-50 dark:bg-slate-950 px-4 py-2.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
          <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 font-mono">Left Copy — Month</label>
          <input type="month" value={monthA} onChange={(e) => setMonthA(e.target.value)}
            className="w-full bg-transparent border-none p-0 m-0 outline-none text-sm font-bold text-slate-900 dark:text-white block appearance-none" />
        </div>
        <div className="bg-slate-50 dark:bg-slate-950 px-4 py-2.5 rounded-2xl border border-slate-200 dark:border-slate-800 focus-within:border-cyan-500 transition-all">
          <label className="block text-[10px] font-bold text-cyan-600 dark:text-cyan-400 uppercase tracking-widest mb-1 font-mono">Right Copy — Month</label>
          <input type="month" value={monthB} onChange={(e) => setMonthB(e.target.value)}
            className="w-full bg-transparent border-none p-0 m-0 outline-none text-sm font-bold text-slate-900 dark:text-white block appearance-none" />
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
          Name, Office, official hours and the verifying officer's details are shared and editable directly on either copy. Each copy prints at half an A4 sheet, side by side, so two months fit on one page — pick a different month per copy if you like.
        </p>
      </div>

      {/* PRINTABLE SHEETS — two half-A4 copies, side by side, centered as a pair */}
      <div className="flex flex-col sm:flex-row print:flex-row justify-center items-stretch gap-4 print:gap-0 max-w-[1000px] mx-auto">
        <div className="sm:w-1/2 print:w-1/2">
          <DTRSheet
            {...sharedFieldProps}
            month={monthA} setMonth={setMonthA}
            rows={rowsA} updateCell={updateCellA}
            dailyTotals={dailyTotalsA} monthTotal={monthTotalA}
            monthLabel={formatMonthLabel(monthA)}
            idPrefix="a"
          />
        </div>
        <div className="hidden print:block w-px border-l border-dashed border-slate-400 mx-2" />
        <div className="sm:w-1/2 print:w-1/2">
          <DTRSheet
            {...sharedFieldProps}
            month={monthB} setMonth={setMonthB}
            rows={rowsB} updateCell={updateCellB}
            dailyTotals={dailyTotalsB} monthTotal={monthTotalB}
            monthLabel={formatMonthLabel(monthB)}
            idPrefix="b"
          />
        </div>
      </div>
    </div>
  );
}
