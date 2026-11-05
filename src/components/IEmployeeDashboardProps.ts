/**
 * Public contract for the Personal Employee Dashboard.
 *
 * The shape mirrors what the SPFx web part class will pass in:
 *  - `user` comes from `this.context.pageContext.user`
 *  - `isDarkTheme` / `quickLinksColumns` are web part properties edited in the property pane
 *  - `quickLinks` / `tasks` are data (static config today, Graph / list data later)
 */

/** Subset of `this.context.pageContext.user` that the dashboard needs. */
export interface IDashboardUser {
  displayName: string;
  email: string;
  /** Optional profile photo, e.g. `/_layouts/15/userphoto.aspx?size=L&accountname=<email>` */
  photoUrl?: string;
  jobTitle?: string;
}

/** Keys of the built-in icon set (see `icons.tsx`). */
export type DashboardIconName =
  | 'mail'
  | 'people'
  | 'receipt'
  | 'calendar'
  | 'chat'
  | 'laptop'
  | 'book'
  | 'heart'
  | 'folder'
  | 'chart';

export interface IQuickLink {
  id: string;
  title: string;
  description?: string;
  url: string;
  icon: DashboardIconName;
  /** Accent color for the icon tile. Any valid CSS color. */
  accent?: string;
  /** Opens in a new tab when true (default: true). */
  openInNewTab?: boolean;
}

export type TaskPriority = 'high' | 'medium' | 'low';

export interface ITask {
  id: string;
  title: string;
  /** Where the task came from, e.g. "HR", "Planner", "Finance". */
  source?: string;
  /** ISO date string (yyyy-mm-dd or full ISO). */
  dueDate?: string;
  priority?: TaskPriority;
  completed: boolean;
  /** ISO date string, set when the task is completed. */
  completedOn?: string;
}

export interface IEmployeeDashboardProps {
  user: IDashboardUser;
  /** Web part property: switches the dashboard to its dark palette. */
  isDarkTheme: boolean;
  /** Web part property: number of columns in the Quick Links grid (clamped to 1–6). */
  quickLinksColumns: number;
  /** Defaults to `DEFAULT_QUICK_LINKS` when omitted. */
  quickLinks?: IQuickLink[];
  tasks?: ITask[];
  /** True while tasks are being fetched (shows a loading placeholder). */
  tasksLoading?: boolean;
  /** Message shown instead of the task list when tasks could not be loaded. */
  tasksError?: string;
  /** Shows a "Try again" button next to `tasksError`. */
  onRetryTasks?: () => void;
  /** Where tasks come from, e.g. "Microsoft To Do". Shown under the heading. */
  tasksSourceName?: string;
  /** Link to the task source, e.g. the To Do web app. */
  tasksSourceUrl?: string;
  /**
   * Called after a task is checked / unchecked, so the host can persist the change.
   * If it returns a promise that rejects, the change is rolled back and an error is shown.
   */
  onTaskToggle?: (task: ITask) => void | Promise<void>;
  /** Override "now" (useful for demos and tests). */
  now?: Date;
}
