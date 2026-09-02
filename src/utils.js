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
