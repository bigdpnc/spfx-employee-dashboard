import * as React from 'react';
import styles from './MyTasks.module.scss';
import { ITask } from './IEmployeeDashboardProps';
import { describeDue, formatShortDate, parseDate } from './dateUtils';
import { ArrowIcon, CheckIcon, ClockIcon } from './icons';

export interface IMyTasksProps {
  tasks: ITask[];
  now: Date;
  onToggle: (id: string) => void;
  loading?: boolean;
  /** Load failure: replaces the list. */
  error?: string;
  onRetry?: () => void;
  /** Save failure for a single task: shown above the list. */
  actionError?: string;
  sourceName?: string;
  sourceUrl?: string;
}

type TabKey = 'pending' | 'completed';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'pending', label: 'Pending' },
  { key: 'completed', label: 'Completed' },
];

const PRIORITY_ORDER = { high: 0, medium: 1, low: 2 };

function sortPending(a: ITask, b: ITask): number {
  // Earliest due first; undated tasks last; then by priority.
  const ad = a.dueDate ? parseDate(a.dueDate).getTime() : Infinity;
  const bd = b.dueDate ? parseDate(b.dueDate).getTime() : Infinity;
  if (ad !== bd) return ad - bd;
  return PRIORITY_ORDER[a.priority || 'medium'] - PRIORITY_ORDER[b.priority || 'medium'];
}

function sortCompleted(a: ITask, b: ITask): number {
  // Most recently completed first.
  const ac = a.completedOn ? parseDate(a.completedOn).getTime() : 0;
  const bc = b.completedOn ? parseDate(b.completedOn).getTime() : 0;
  return bc - ac;
}

export const MyTasks: React.FC<IMyTasksProps> = ({
  tasks,
  now,
  onToggle,
  loading = false,
  error,
  onRetry,
  actionError,
  sourceName,
  sourceUrl,
}) => {
  const [activeTab, setActiveTab] = React.useState<TabKey>('pending');
  const tabRefs = React.useRef<Record<TabKey, HTMLButtonElement | null>>({ pending: null, completed: null });

  const pending = React.useMemo(() => tasks.filter((t) => !t.completed).sort(sortPending), [tasks]);
  const completed = React.useMemo(() => tasks.filter((t) => t.completed).sort(sortCompleted), [tasks]);
  const visible = activeTab === 'pending' ? pending : completed;
  const counts: Record<TabKey, number> = { pending: pending.length, completed: completed.length };

  // WAI-ARIA tabs pattern: arrow keys / Home / End move between tabs.
  const onTabKeyDown = (e: React.KeyboardEvent<HTMLButtonElement>): void => {
    const idx = TABS.findIndex((t) => t.key === activeTab);
    let next = idx;
    if (e.key === 'ArrowRight') next = (idx + 1) % TABS.length;
    else if (e.key === 'ArrowLeft') next = (idx - 1 + TABS.length) % TABS.length;
    else if (e.key === 'Home') next = 0;
    else if (e.key === 'End') next = TABS.length - 1;
    else return;
    e.preventDefault();
    const key = TABS[next].key;
    setActiveTab(key);
    const el = tabRefs.current[key];
    if (el) el.focus();
  };

  const progress = tasks.length === 0 ? 0 : Math.round((completed.length / tasks.length) * 100);

  return (
    <section className={styles.card} aria-labelledby="ed-tasks-heading">
      <div className={styles.top}>
        <div>
          <h3 id="ed-tasks-heading">My Tasks</h3>
          <p className={styles.progressLabel}>
            {loading || error ? null : (
              <span>
                {completed.length} of {tasks.length} done
              </span>
            )}
            {sourceName && sourceUrl ? (
              <a className={styles.sourceLink} href={sourceUrl} target="_blank" rel="noopener noreferrer">
                {sourceName}
                <ArrowIcon size={12} />
                <span className={styles.srOnly}> (opens in a new tab)</span>
              </a>
            ) : sourceName ? (
              <span>{sourceName}</span>
            ) : null}
          </p>
        </div>
        <div
          className={styles.progress}
          role="progressbar"
          aria-label="Task completion"
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={progress}
        >
          <span style={{ width: `${progress}%` }} />
        </div>
      </div>

      <div className={styles.tabs} role="tablist" aria-label="Task status">
        {TABS.map((tab) => {
          const selected = tab.key === activeTab;
          return (
            <button
              key={tab.key}
              ref={(el) => {
                tabRefs.current[tab.key] = el;
              }}
              type="button"
              role="tab"
              id={`ed-tab-${tab.key}`}
              aria-selected={selected}
              aria-controls={`ed-tabpanel-${tab.key}`}
              tabIndex={selected ? 0 : -1}
              className={`${styles.tab} ${selected ? styles.tabActive : ''}`}
              onClick={() => setActiveTab(tab.key)}
              onKeyDown={onTabKeyDown}
            >
              {tab.label}
              {loading || error ? null : <span className={styles.badge}>{counts[tab.key]}</span>}
            </button>
          );
        })}
      </div>

      <div
        className={styles.panel}
        role="tabpanel"
        id={`ed-tabpanel-${activeTab}`}
        aria-labelledby={`ed-tab-${activeTab}`}
        tabIndex={0}
      >
        {actionError && !loading && !error ? (
          <p className={styles.actionError} role="alert">
            {actionError}
          </p>
        ) : null}
        {loading ? (
          <ul className={styles.list} aria-busy="true" aria-label="Loading tasks">
            {[0, 1, 2].map((i) => (
              <li key={i} className={styles.item}>
                <span className={`${styles.skeleton} ${styles.skeletonBox}`} />
                <span className={styles.itemBody}>
                  <span className={`${styles.skeleton} ${styles.skeletonLine}`} style={{ width: `${70 - i * 12}%` }} />
                  <span className={`${styles.skeleton} ${styles.skeletonLine} ${styles.skeletonShort}`} />
                </span>
              </li>
            ))}
          </ul>
        ) : error ? (
          <div className={styles.errorState} role="alert">
            <p>{error}</p>
            {onRetry ? (
              <button type="button" className={styles.retry} onClick={onRetry}>
                Try again
              </button>
            ) : null}
          </div>
        ) : visible.length === 0 ? (
          <div className={styles.empty}>
            <span className={styles.emptyIcon} aria-hidden="true">
              <CheckIcon size={22} />
            </span>
            <p>{activeTab === 'pending' ? "Nothing pending. You're all caught up!" : 'No completed tasks yet.'}</p>
          </div>
        ) : (
          <ul className={styles.list}>
            {visible.map((task) => {
              const due = !task.completed && task.dueDate ? describeDue(task.dueDate, now) : null;
              return (
                <li key={task.id} className={`${styles.item} ${task.completed ? styles.done : ''}`}>
                  <button
                    type="button"
                    className={styles.check}
                    role="checkbox"
                    aria-checked={task.completed}
                    aria-label={`${task.completed ? 'Mark as pending' : 'Mark as complete'}: ${task.title}`}
                    onClick={() => onToggle(task.id)}
                  >
                    <CheckIcon size={14} />
                  </button>
                  <div className={styles.itemBody}>
                    <span className={styles.itemTitle}>{task.title}</span>
                    <span className={styles.meta}>
                      {task.source ? <span className={styles.source}>{task.source}</span> : null}
                      {due ? (
                        <span className={`${styles.due} ${styles[due.tone]}`}>
                          <ClockIcon size={13} />
                          {due.label}
                        </span>
                      ) : null}
                      {task.completed && task.completedOn ? (
                        <span className={styles.due}>Completed {formatShortDate(task.completedOn)}</span>
                      ) : null}
                    </span>
                  </div>
                  {!task.completed && task.priority ? (
                    <span className={`${styles.priority} ${styles[task.priority]}`}>{task.priority}</span>
                  ) : null}
                </li>
              );
            })}
          </ul>
        )}
      </div>
    </section>
  );
};
