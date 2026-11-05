import { IQuickLink } from './IEmployeeDashboardProps';

/** Used when the host doesn't pass its own `quickLinks`. Swap URLs for your tenant. */
export const DEFAULT_QUICK_LINKS: IQuickLink[] = [
  {
    id: 'email',
    title: 'Email',
    description: 'Open Outlook on the web',
    url: 'https://outlook.office.com/mail/',
    icon: 'mail',
    accent: '#2563eb',
  },
  {
    id: 'hr',
    title: 'HR Portal',
    description: 'Time off, payslips and profile',
    url: '#hr-portal',
    icon: 'people',
    accent: '#7c3aed',
  },
  {
    id: 'expenses',
    title: 'Expense Reports',
    description: 'Submit and track claims',
    url: '#expenses',
    icon: 'receipt',
    accent: '#059669',
  },
  {
    id: 'calendar',
    title: 'Calendar',
    description: 'Meetings and room booking',
    url: 'https://outlook.office.com/calendar/',
    icon: 'calendar',
    accent: '#ea580c',
  },
  {
    id: 'teams',
    title: 'Teams',
    description: 'Chats, channels and calls',
    url: 'https://teams.microsoft.com/',
    icon: 'chat',
    accent: '#4f46e5',
  },
  {
    id: 'it',
    title: 'IT Help Desk',
    description: 'Raise or follow a ticket',
    url: '#it-help',
    icon: 'laptop',
    accent: '#0891b2',
  },
  {
    id: 'learning',
    title: 'Learning',
    description: 'Courses and certifications',
    url: '#learning',
    icon: 'book',
    accent: '#db2777',
  },
  {
    id: 'benefits',
    title: 'Benefits',
    description: 'Health, wellbeing and perks',
    url: '#benefits',
    icon: 'heart',
    accent: '#dc2626',
  },
];
