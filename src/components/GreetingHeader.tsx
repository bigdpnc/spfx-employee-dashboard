import * as React from 'react';
import styles from './GreetingHeader.module.scss';
import { IDashboardUser } from './IEmployeeDashboardProps';
import { formatLongDate, getGreeting } from './dateUtils';
import { SparkleIcon } from './icons';

export interface IGreetingHeaderProps {
  user: IDashboardUser;
  now: Date;
  pendingCount: number;
  dueTodayCount: number;
  overdueCount: number;
  /** False while tasks are loading or failed to load, so the summary doesn't claim "all caught up". */
  tasksAvailable?: boolean;
}

function getInitials(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  const first = parts[0][0];
  const last = parts.length > 1 ? parts[parts.length - 1][0] : '';
  return (first + last).toUpperCase();
}

/** "Megan Bowen" → "Megan"; "Bowen, Megan" → "Megan". */
function getFirstName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.indexOf(',') > -1) return trimmed.split(',')[1].trim().split(/\s+/)[0] || trimmed;
  return trimmed.split(/\s+/)[0] || 'there';
}

function buildSummary(pending: number, dueToday: number, overdue: number): string {
  if (pending === 0) return "You're all caught up. Nice work!";
  const parts = [`You have ${pending} pending task${pending === 1 ? '' : 's'}`];
  if (overdue > 0) parts.push(`${overdue} overdue`);
  if (dueToday > 0) parts.push(`${dueToday} due today`);
  return parts.join(' · ');
}

export const GreetingHeader: React.FC<IGreetingHeaderProps> = ({
  user,
  now,
  pendingCount,
  dueTodayCount,
  overdueCount,
  tasksAvailable = true,
}) => {
  const [photoFailed, setPhotoFailed] = React.useState(false);
  React.useEffect(() => setPhotoFailed(false), [user.photoUrl]);

  const showPhoto = !!user.photoUrl && !photoFailed;

  return (
    <header className={styles.header}>
      <div className={styles.glow} aria-hidden="true" />
      <div className={styles.avatar} aria-hidden="true">
        {showPhoto ? (
          <img src={user.photoUrl} alt="" onError={() => setPhotoFailed(true)} />
        ) : (
          <span>{getInitials(user.displayName)}</span>
        )}
      </div>
      <div className={styles.text}>
        <p className={styles.date}>
          <SparkleIcon size={14} />
          {formatLongDate(now)}
        </p>
        <h2 className={styles.title}>
          {getGreeting(now)}, {getFirstName(user.displayName)}
        </h2>
        <p className={styles.summary}>
          {tasksAvailable ? buildSummary(pendingCount, dueTodayCount, overdueCount) : "Here's your day at a glance."}
          {user.jobTitle ? <span className={styles.role}>{user.jobTitle}</span> : null}
        </p>
      </div>
    </header>
  );
};
