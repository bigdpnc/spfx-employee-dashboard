/** Date helpers kept dependency-free (no moment / date-fns needed in the bundle). */

export function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** Parses `yyyy-mm-dd` as a *local* date (the native parser treats it as UTC). */
export function parseDate(value: string): Date {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date(value);
}

/** Whole days from `now` to `value` (negative = in the past). */
export function daysFrom(value: string, now: Date): number {
  const ms = startOfDay(parseDate(value)).getTime() - startOfDay(now).getTime();
  return Math.round(ms / 86400000);
}

export function getGreeting(now: Date): string {
  const h = now.getHours();
  if (h < 5) return 'Working late';
  if (h < 12) return 'Good morning';
  if (h < 17) return 'Good afternoon';
  return 'Good evening';
}

export function formatLongDate(now: Date): string {
  return now.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
}

export function formatShortDate(value: string): string {
  return parseDate(value).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
}

/** "Overdue by 2 days", "Due today", "Due tomorrow", "Due Nov 12". */
export function describeDue(value: string, now: Date): { label: string; tone: 'overdue' | 'soon' | 'normal' } {
  const d = daysFrom(value, now);
  if (d < 0) return { label: `Overdue by ${-d} day${d === -1 ? '' : 's'}`, tone: 'overdue' };
  if (d === 0) return { label: 'Due today', tone: 'soon' };
  if (d === 1) return { label: 'Due tomorrow', tone: 'soon' };
  return { label: `Due ${formatShortDate(value)}`, tone: 'normal' };
}
