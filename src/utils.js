export function calculateHoursRendered(inStr, outStr, shouldDeduct) {
  const start = new Date(`1970-01-01T${inStr}:00Z`);
  const end = new Date(`1970-01-01T${outStr}:00Z`);
  let diff = (end - start) / 1000 / 60 / 60; 
  if (shouldDeduct && diff > 5) diff -= 1;
  return diff > 0 ? diff.toFixed(2) : 0;
}

export function getCurrentTime() {
  const now = new Date();
  return `${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
}

export function getGreeting() {
  const hour = new Date().getHours();
  if (hour < 12) return 'Good Morning';
  if (hour < 18) return 'Good Afternoon';
  return 'Good Evening';
}

export function formatDateString(dateStr) {
  return new Date(dateStr).toLocaleDateString('en-US', { weekday: 'short', month: 'short', day: 'numeric', year: 'numeric' });
}

export function getInitials(name) {
  if (!name) return 'S';
  return name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
}

export function getHoursDiff(start, end) {
  if (!start || !end) return 0;
  const [sh, sm] = start.split(':').map(Number);
  const [eh, em] = end.split(':').map(Number);
  const diff = (eh * 60 + em - (sh * 60 + sm)) / 60;
  return diff > 0 ? diff : 0;
}

export function getDaysInMonth(monthStr) {
  const [y, m] = (monthStr || '').split('-').map(Number);
  if (!y || !m) return 31;
  return new Date(y, m, 0).getDate();
}

export function formatMonthLabel(monthStr) {
  const [y, m] = (monthStr || '').split('-').map(Number);
  if (!y || !m) return '';
  return new Date(y, m - 1, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
}

export function shiftMonth(monthStr, delta) {
  const [y, m] = (monthStr || '').split('-').map(Number);
  if (!y || !m) return monthStr;
  const d = new Date(y, m - 1 + delta, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
