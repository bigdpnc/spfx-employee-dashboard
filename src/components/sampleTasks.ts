import { ITask } from './IEmployeeDashboardProps';

/**
 * Example tasks shown until a real task source (e.g. Microsoft To Do or Planner via Graph)
 * is connected. Dates are relative to today so the due/overdue states always show.
 */

/** yyyy-mm-dd for `offset` days from today, so the demo data never goes stale. */
function day(offset: number): string {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export const SAMPLE_TASKS: ITask[] = [
  { id: 't1', title: 'Submit Q4 travel expense report', source: 'Finance', dueDate: day(-1), priority: 'high', completed: false },
  { id: 't2', title: 'Complete annual security awareness training', source: 'Learning', dueDate: day(0), priority: 'high', completed: false },
  { id: 't3', title: 'Review design system tokens with engineering', source: 'Planner', dueDate: day(1), priority: 'medium', completed: false },
  { id: 't4', title: 'Confirm open enrollment benefit choices', source: 'HR', dueDate: day(6), priority: 'medium', completed: false },
  { id: 't5', title: 'Book a desk for the team offsite', source: 'Facilities', priority: 'low', completed: false },
  { id: 't6', title: 'Update emergency contact details', source: 'HR', completed: true, completedOn: day(-1) },
  { id: 't7', title: 'Approve timesheet for last week', source: 'HR', completed: true, completedOn: day(-3) },
  { id: 't8', title: 'Share onboarding feedback survey', source: 'People Ops', completed: true, completedOn: day(-5) },
];
