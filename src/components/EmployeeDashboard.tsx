import * as React from 'react';
import styles from './EmployeeDashboard.module.scss';
import { IEmployeeDashboardProps, ITask } from './IEmployeeDashboardProps';
import { GreetingHeader } from './GreetingHeader';
import { QuickLinks } from './QuickLinks';
import { MyTasks } from './MyTasks';
import { DEFAULT_QUICK_LINKS } from './defaultQuickLinks';
import { daysFrom } from './dateUtils';

const EMPTY_TASKS: ITask[] = [];

/** Re-renders once a minute so the greeting rolls over (morning → afternoon) on a long-open page. */
function useNow(override?: Date): Date {
  const [now, setNow] = React.useState(() => override || new Date());
  React.useEffect(() => {
    if (override) {
      setNow(override);
      return undefined;
    }
    const id = window.setInterval(() => setNow(new Date()), 60000);
    return () => window.clearInterval(id);
  }, [override]);
  return now;
}

function toIsoDate(d: Date): string {
  const pad = (n: number): string => (n < 10 ? `0${n}` : `${n}`);
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export const EmployeeDashboard: React.FC<IEmployeeDashboardProps> = (props) => {
  const { user, isDarkTheme, quickLinksColumns, quickLinks = DEFAULT_QUICK_LINKS, tasks: taskProp = EMPTY_TASKS, onTaskToggle } = props;
  const now = useNow(props.now);

  // Local copy so checking a task is instant; the host persists via onTaskToggle.
  const [tasks, setTasks] = React.useState<ITask[]>(taskProp);
  React.useEffect(() => setTasks(taskProp), [taskProp]);

  const handleToggle = React.useCallback(
    (id: string) => {
      const current = tasks.find((t) => t.id === id);
      if (!current) return;
      const updated: ITask = current.completed
        ? { ...current, completed: false, completedOn: undefined }
        : { ...current, completed: true, completedOn: toIsoDate(new Date()) };
      setTasks((prev) => prev.map((t) => (t.id === id ? updated : t)));
      if (onTaskToggle) onTaskToggle(updated);
    },
    [tasks, onTaskToggle],
  );

  const pending = tasks.filter((t) => !t.completed);
  const overdueCount = pending.filter((t) => t.dueDate && daysFrom(t.dueDate, now) < 0).length;
  const dueTodayCount = pending.filter((t) => t.dueDate && daysFrom(t.dueDate, now) === 0).length;

  return (
    <div className={`${styles.dashboard} ${isDarkTheme ? styles.dark : styles.light}`} data-theme={isDarkTheme ? 'dark' : 'light'}>
      <div className={styles.inner}>
        <GreetingHeader
          user={user}
          now={now}
          pendingCount={pending.length}
          dueTodayCount={dueTodayCount}
          overdueCount={overdueCount}
        />
        <QuickLinks links={quickLinks} columns={quickLinksColumns} />
        <MyTasks tasks={tasks} now={now} onToggle={handleToggle} />
      </div>
    </div>
  );
};

export default EmployeeDashboard;
